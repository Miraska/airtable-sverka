require('dotenv').config(); // Если используете .env-файл
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');

// Импортируем S3-клиент и утилиты для presigned URL
const {
  S3Client,
  PutObjectCommand,
  GetObjectCommand
} = require('@aws-sdk/client-s3');
const { getSignedUrl } = require('@aws-sdk/s3-request-presigner');

const app = express();

// Разрешаем CORS (если нужно обращаться извне)
app.use(cors());

// Для парсинга JSON-тел запросов
app.use(bodyParser.json());

const PORT = 3000;

// Путь к локальному шаблону (используем, если в S3 ещё нет файла)
const excelFilePath = path.join(__dirname, 'template.xlsx');

// Проверяем, что локальный шаблон существует
if (!fs.existsSync(excelFilePath)) {
  console.error('Excel-файл не найден по пути:', excelFilePath);
  process.exit(1);
}

// Инициализируем S3-клиент для Яндекс Облака
const s3Client = new S3Client({
  region: 'ru-central1', // регион Яндекс Облака
  endpoint: 'https://storage.yandexcloud.net',
  credentials: {
    accessKeyId: process.env.S3_ACCESS_KEY,
    secretAccessKey: process.env.S3_SECRET_KEY
  }
});

// Имя бакета берем из .env (или можно захардкодить)
const bucketName = process.env.S3_BUCKET || 'my-bucket-name';

// Ключ (имя) файла в бакете, где будем хранить обновлённую версию
// (Можно назвать как угодно, главное всегда читать/записывать один и тот же Key)
const s3Key = 'uploads/updatedFile-latest.xlsx';

// Тестовый GET
app.get('/', (req, res) => {
  res.send('Hello World! Спробуйте POST /receive-data');
});

/**
 * Функция: Загрузка "актуального" файла Excel из S3, если он там есть.
 * Если файла нет, возвращаем null — тогда будем использовать локальный шаблон.
 */
async function loadWorkbookFromS3() {
  try {
    const response = await s3Client.send(
      new GetObjectCommand({
        Bucket: bucketName,
        Key: s3Key
      })
    );
    // response.Body — это поток (stream)
    const stream = response.Body;
    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    // Превращаем buffer в объект XlsxPopulate
    return XlsxPopulate.fromDataAsync(buffer);
  } catch (error) {
    if (error.name === 'NoSuchKey' || error.$metadata?.httpStatusCode === 404) {
      // Файл в S3 не найден
      return null;
    }
    throw error;
  }
}

/**
 * Функция: Загружает buffer (содержимое Excel) в S3.
 * Мы всегда перезаписываем один и тот же Key (s3Key).
 */
async function uploadWorkbookToS3(buffer) {
  await s3Client.send(
    new PutObjectCommand({
      Bucket: bucketName,
      Key: s3Key,
      Body: buffer,
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })
  );
}

/**
 * Функция: находим первую свободную строку, начиная с заданной (startRow).
 * Проверяем ячейки в колонке B. Если она занята — идём вниз.
 * Возвращаем номер строки, которая оказалась пустой.
 */
function findNextEmptyRow(sheet, startRow = 3) {
  let row = startRow;
  // Пока в B<row> что-то есть — идём дальше.
  while (sheet.cell(`B${row}`).value()) {
    row++;
  }
  return row;
}

/**
 * POST-запрос:
 * 1) Получаем JSON-данные
 * 2) Пытаемся загрузить уже существующий файл из S3 (если он есть)
 *    либо берём локальный template.xlsx
 * 3) Находим первую свободную строку
 * 4) Записываем новые данные
 * 5) Генерируем итоговый Excel-файл в буфере
 * 6) Перезаписываем файл в S3
 * 7) Генерируем presigned URL и возвращаем клиенту
 */
app.post('/receive-data', async (req, res) => {
  console.log('Получен POST-запрос на /receive-data');
  try {
    // Данные, которые прислали (например, из Airtable)
    const data = req.body;
    console.log('Получены данные:', data);

    // 1) Пытаемся загрузить уже обновлённый файл из S3
    let workbook = await loadWorkbookFromS3();

    // 2) Если в S3 файла нет, открываем локальный шаблон
    if (!workbook) {
      console.log('Файл в S3 не найден, используем локальный template.xlsx');
      workbook = await XlsxPopulate.fromFileAsync(excelFilePath);
    }

    // Берём первый лист (или нужный вам индекс)
    const sheet = workbook.sheet(0);

    // 3) Находим первую пустую строку, начиная с 3-й
    const targetRow = findNextEmptyRow(sheet, 3);

    // 4) Записываем данные в нужные ячейки
    sheet.cell(`B${targetRow}`).value(data.Дата || '');
    sheet.cell(`D${targetRow}`).value(data.Контрагент?.[0]?.name || '');
    sheet.cell(`E${targetRow}`).value(data.Получатель?.[0]?.name || '');
    sheet.cell(`J${targetRow}`).value(data['Сумма USD'] || 0);

    // 5) Генерируем новый Excel-файл в буфере
    const buffer = await workbook.outputAsync();

    // 6) Загружаем (перезаписываем) обновлённый файл в S3
    await uploadWorkbookToS3(buffer);
    console.log(`Файл перезаписан в S3: s3://${bucketName}/${s3Key}`);

    // 7) Генерируем presigned URL (например, на 1 час)
    const downloadUrl = await getSignedUrl(
      s3Client,
      new GetObjectCommand({
        Bucket: bucketName,
        Key: s3Key
      }),
      { expiresIn: 3600 } // 1 час
    );

    // Возвращаем ссылку
    return res.json({
      success: true,
      message: 'File created/updated and uploaded successfully',
      fileUrl: downloadUrl
    });
  } catch (error) {
    console.error('Ошибка при обработке данных:', error);
    res.status(500).json({
      success: false,
      message: 'Ошибка при обработке данных'
    });
  }
});

/**
 * Пример GET-запроса на скачивание локального template.xlsx (для отладки).
 * Можно не использовать, если не нужно.
 */
app.get('/download-latest', (req, res) => {
  const file = path.join(__dirname, 'template.xlsx');
  res.download(file, 'Updated.xlsx', err => {
    if (err) {
      console.error('Ошибка при скачивании:', err);
      res.status(500).send('Ошибка при скачивании файла');
    }
  });
});

// Запуск сервера
app.listen(PORT, () => {
  console.log(`Сервер запущен на порту ${PORT}`);
  console.log(`POST на http://localhost:${PORT}/receive-data`);
});
