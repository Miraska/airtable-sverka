require("dotenv").config(); // Если используете .env-файл
const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const XlsxPopulate = require("xlsx-populate");

// Импортируем S3-клиент и утилиты для presigned URL
const {
  S3Client,
  PutObjectCommand,
  GetObjectCommand,
} = require("@aws-sdk/client-s3");
const { getSignedUrl } = require("@aws-sdk/s3-request-presigner");

const app = express();

app.use(cors());
app.use(bodyParser.json());

const PORT = 3000;
let fileName;

// Путь к локальному шаблону
const excelFilePath = path.join(__dirname, "template.xlsx");
if (!fs.existsSync(excelFilePath)) {
  console.error("Excel-файл не найден по пути:", excelFilePath);
  process.exit(1);
}

// Инициализируем S3-клиент (пример для Яндекс Облака)
const s3Client = new S3Client({
  region: "ru-central1",
  endpoint: "https://storage.yandexcloud.net",
  credentials: {
    accessKeyId: process.env.S3_ACCESS_KEY,
    secretAccessKey: process.env.S3_SECRET_KEY,
  },
});

const bucketName = process.env.S3_BUCKET || "my-bucket-name";

app.get("/", (req, res) => {
  res.send("Hello World! Попробуйте POST /receive-data");
});

/**
 * Пытаемся загрузить уже существующий файл из S3 (если есть).
 * Если нет — вернётся null, тогда возьмём локальный шаблон.
 */
async function loadWorkbookFromS3() {
  try {
    const response = await s3Client.send(
      new GetObjectCommand({
        Bucket: bucketName,
        Key: `uploads/${fileName}.xlsx`,
      })
    );
    const stream = response.Body;
    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    return XlsxPopulate.fromDataAsync(buffer);
  } catch (error) {
    if (error.name === "NoSuchKey" || error.$metadata?.httpStatusCode === 404) {
      return null; // Файл не найден
    }
    throw error;
  }
}

/** Загрузка (перезапись) книги в S3 */
async function uploadWorkbookToS3(buffer) {
  await s3Client.send(
    new PutObjectCommand({
      Bucket: bucketName,
      Key: `uploads/${fileName}.xlsx`,
      Body: buffer,
      ContentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    })
  );
}

/**
 * Поиск первой свободной строки, начиная с startRow,
 * проверяем ячейку в колонке B (например).
 */
function findNextEmptyRow(sheet, startRow = 3) {
  let row = startRow;
  while (sheet.cell(`B${row}`).value()) {
    row++;
  }
  return row;
}



/**
 * Функция заполнения одной строки: данные в нужные столбцы и формулы
 * для "промежуточного" баланса, исключая строки "Итого".
 *
 * @param {Object} sheet - Лист XlsxPopulate
 * @param {number} rowIndex - Индекс строки, которую заполняем
 * @param {Object} record - Объект с данными (например, из JSON)
 * @param {number} [startRow=3] - С какой строки начинаются записи
 */
function fillRow(sheet, rowIndex, record, startRow = 3) {
  // Колонка A — для примера вставляем "1", 
  // у вас может быть другая логика (или пусто)
  sheet.cell(`A${rowIndex}`).value(1);

  // Преобразуем дату к формату дд.мм.гггг
  // Если значение в record.Дата корректное, парсим и ставим формат
  const dateObj = new Date(record.Дата);
  if (!isNaN(dateObj.valueOf())) {
    sheet.cell(`B${rowIndex}`).value(dateObj);
    // Устанавливаем нужный формат даты
    sheet.cell(`B${rowIndex}`).style("numberFormat", "dd.mm.yyyy");
  } else {
    // Если не дата, вставляем как есть
    sheet.cell(`B${rowIndex}`).value(record.Дата || "");
  }

  // Плательщик (у вас в JSON это "Отправитель")
  sheet.cell(`D${rowIndex}`).value(record.Отправитель?.[0]?.name || "");

  // Получатель
  sheet.cell(`E${rowIndex}`).value(record.Получатель?.[0]?.name || "");

  // Курс (может быть массивом или строкой)
  const course = Array.isArray(record["Курс"])
    ? record["Курс"].join("; ")
    : record["Курс"] || "";
  sheet.cell(`F${rowIndex}`).value(course);

  // Комментарий (у вас в JSON это "Комментарий (from Ордер)")
  const comment = Array.isArray(record["Комментарий (from Ордер)"])
    ? record["Комментарий (from Ордер)"].join("; ")
    : record["Комментарий (from Ордер)"] || "";
  sheet.cell(`AA${rowIndex}`).value(comment);

  // ----- Раскладка по валютам -----
  // Рубли
  if (record["Сумма RUB"]) {
    // К выдаче рубли - колонка H
    sheet.cell(`H${rowIndex}`).value(record["Сумма RUB"]);
  }

  // Доллары
  if (record["Сумма USD"]) {
    // К выдаче доллары → колонка J
    sheet.cell(`J${rowIndex}`).value(record["Сумма USD"]);
  }

  // USDT
  if (record["Сумма USDT"]) {
    // К выдаче USDT (Тезер) → колонка L
    sheet.cell(`L${rowIndex}`).value(record["Сумма USDT"]);
  }

  // EURO
  if (record["Сумма EURO"]) {
    sheet.cell(`N${rowIndex}`).value(record["Сумма EURO"]);
  }

  // CNY
  if (record["Сумма CNY"]) {
    sheet.cell(`P${rowIndex}`).value(record["Сумма CNY"]);
  }

  // AED
  if (record["Сумма AED"]) {
    sheet.cell(`R${rowIndex}`).value(record["Сумма AED"]);
  }

  // ----- Формулы для промежуточного баланса -----
  // Используем SUMIF, чтобы не учитывались строки, где в A = "Итого".

  // Баланс на конец дня (Рубли = H + I)
  sheet.cell(`U${rowIndex}`).formula(
    `=SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", H${startRow}:H${rowIndex})` +
    `+SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", I${startRow}:I${rowIndex})`
  );

  // Баланс на конец дня (Доллары = J + K)
  sheet.cell(`V${rowIndex}`).formula(
    `=SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", J${startRow}:J${rowIndex})` +
    `+SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", K${startRow}:K${rowIndex})`
  );

  // Баланс на конец дня (USDT = L + M)
  sheet.cell(`W${rowIndex}`).formula(
    `=SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", L${startRow}:L${rowIndex})` +
    `+SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", M${startRow}:M${rowIndex})`
  );

  // Баланс на конец дня (EURO = N + O)
  sheet.cell(`X${rowIndex}`).formula(
    `=SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", N${startRow}:N${rowIndex})` +
    `+SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", O${startRow}:O${rowIndex})`
  );

  // Баланс на конец дня (CNY = P + Q)
  sheet.cell(`Y${rowIndex}`).formula(
    `=SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", P${startRow}:P${rowIndex})` +
    `+SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", Q${startRow}:Q${rowIndex})`
  );

  // Баланс на конец дня (AED = R + S)
  sheet.cell(`Z${rowIndex}`).formula(
    `=SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", R${startRow}:R${rowIndex})` +
    `+SUMIF($A$${startRow}:$A$${rowIndex}, "<>Итого", S${startRow}:S${rowIndex})`
  );
}


/**
 * Вставляем строку «итоговой» формулы и заливаем её чёрным цветом
 * (текст делаем белым). Формулы для столбцов H..O (складываем значения).
 */
function addSummaryRow(sheet, rowIndex, startRow, endRow) {
  // Формулы для валют
  sheet.cell(`H${rowIndex}`).formula(`=SUM(H${startRow}:H${endRow})`);
  sheet.cell(`I${rowIndex}`).formula(`=SUM(I${startRow}:I${endRow})`);
  sheet.cell(`J${rowIndex}`).formula(`=SUM(J${startRow}:J${endRow})`);
  sheet.cell(`K${rowIndex}`).formula(`=SUM(K${startRow}:K${endRow})`);
  sheet.cell(`L${rowIndex}`).formula(`=SUM(L${startRow}:L${endRow})`);
  sheet.cell(`M${rowIndex}`).formula(`=SUM(M${startRow}:M${endRow})`);
  sheet.cell(`N${rowIndex}`).formula(`=SUM(N${startRow}:N${endRow})`);
  sheet.cell(`O${rowIndex}`).formula(`=SUM(O${startRow}:O${endRow})`);
  sheet.cell(`P${rowIndex}`).formula(`=SUM(P${startRow}:P${endRow})`);
  sheet.cell(`Q${rowIndex}`).formula(`=SUM(Q${startRow}:Q${endRow})`);
  sheet.cell(`R${rowIndex}`).formula(`=SUM(R${startRow}:R${endRow})`);
  sheet.cell(`S${rowIndex}`).formula(`=SUM(S${startRow}:S${endRow})`);

  sheet.cell(`U${rowIndex}`).formula(`=H${rowIndex}`);
  sheet.cell(`V${rowIndex}`).formula(`=J${rowIndex}`);
  sheet.cell(`W${rowIndex}`).formula(`=L${rowIndex}`);
  sheet.cell(`X${rowIndex}`).formula(`=N${rowIndex}`);
  sheet.cell(`Y${rowIndex}`).formula(`=P${rowIndex}`);
  sheet.cell(`Z${rowIndex}`).formula(`=R${rowIndex}`);

  // Заливаем всю строку чёрным, делаем шрифт белым
  sheet.row(rowIndex).style({
    fill: "808080",
    fontColor: "FFFFFF",
    bold: true,
  });
}

/**
 * POST /receive-data
 * 1) Принимаем массив данных
 * 2) Грузим/создаём Excel
 * 3) Сортируем по дате
 * 4) Группируем (вставляем серую/итоговую строку при смене даты)
 * 5) Сохраняем в S3
 * 6) Отдаём presigned URL
 */
app.post("/receive-data", async (req, res) => {
  console.log("Получен POST-запрос на /receive-data");
  try {
    fileName = req.body.fileName; // Имя файла сразу сохраянем

    let data = req.body.data; // Массив записей
    console.log("Получены данные:\n", JSON.stringify(data, null, 2));

    // 1) Пытаемся взять актуальный файл из S3
    // let workbook = await loadWorkbookFromS3();

    // Если нет в S3 - берём локальный шаблон
    // if (!workbook) {
    //   console.log('Файл в S3 не найден, используем локальный template.xlsx');
    //   workbook = await XlsxPopulate.fromFileAsync(excelFilePath);
    // }

    // Для наглядности здесь просто сразу берем локальный template.xlsx
    console.log("Файл в S3 не найден, используем локальный template.xlsx");
    let workbook = await XlsxPopulate.fromFileAsync(excelFilePath);

    const sheet = workbook.sheet(0);

    // 2) Сортируем записи по дате (если нужно — дополнительно по времени)
    data.sort((a, b) => {
      return new Date(a.Дата) - new Date(b.Дата);
    });

    // Вспомогательная функция для вставки "итоговой" строки
    function insertSummaryRow() {
      sheet.cell(`A${currentRow}`).value("Итого:");
      addSummaryRow(sheet, currentRow, groupStartRow, currentRow - 1);
      currentRow++;
    }

    // 3) Находим первую пустую строку в Excel
    let currentRow = findNextEmptyRow(sheet, 3);

    // Запомним, где начинается "первая группа" (для сумм)
    let groupStartRow = currentRow;
    let lastDate = null;

    // Перебираем записи
    for (const record of data) {
      const currentDate = record.Дата || "";

      // Если дата изменилась и у нас уже была дата, вставляем итог
      if (lastDate && lastDate !== currentDate) {
        insertSummaryRow();
        groupStartRow = currentRow;
      }

      lastDate = currentDate;

      // Заполняем текущую строку данными
      fillRow(sheet, currentRow, record);

      // Переходим на следующую строку
      currentRow++;
    }

    console.log("Всего строк:", currentRow);

    // Когда записи закончились, но осталась "последняя" группа
    if (data.length > 0) {
      insertSummaryRow();
    }

    // При желании можно скрыть неиспользуемые столбцы P..T, если нужно:
    // for (let col = 16; col <= 20; col++) {
    //   sheet.column(col).hidden(true);
    // }

    // 4) Генерируем буфер и сохраняем обновлённый Excel обратно в S3
    const buffer = await workbook.outputAsync();
    await uploadWorkbookToS3(buffer);

    // 5) Генерируем presigned URL на скачивание
    const downloadUrl = await getSignedUrl(
      s3Client,
      new GetObjectCommand({
        Bucket: bucketName,
        Key: `uploads/${fileName}.xlsx`,
      }),
      { expiresIn: 3600 }
    );

    return res.json({
      success: true,
      message: "Файл успешно обновлён и загружен в S3",
      fileUrl: downloadUrl,
    });
  } catch (error) {
    console.error("Ошибка при обработке данных:", error);
    res.status(500).json({
      success: false,
      message: "Ошибка при обработке данных",
    });
  }
});

app.get("/download-latest", (req, res) => {
  const file = path.join(__dirname, "template.xlsx");
  res.download(file, "Updated.xlsx", (err) => {
    if (err) {
      console.error("Ошибка при скачивании:", err);
      res.status(500).send("Ошибка при скачивании файла");
    }
  });
});

app.listen(PORT, () => {
  console.log(`Сервер запущен на порту ${PORT}`);
  console.log(`POST на http://localhost:${PORT}/receive-data`);
});
