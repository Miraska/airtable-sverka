require('dotenv').config(); // Если используете .env-файл
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');

// Импортируем S3-клиент и утилиты для presigned URL
const { S3Client, PutObjectCommand, GetObjectCommand } = require('@aws-sdk/client-s3');
const { getSignedUrl } = require('@aws-sdk/s3-request-presigner');

const app = express();

// Разрешаем CORS (если нужно обращаться извне)
app.use(cors());

// Для парсинга JSON-тел запросов
app.use(bodyParser.json());

const PORT = 3000;

// Путь к существующему Excel-шаблону (со стилями)
const excelFilePath = path.join(__dirname, 'template.xlsx');

// Проверяем, что файл существует
if (!fs.existsSync(excelFilePath)) {
    console.error('Excel-файл не найден по пути:', excelFilePath);
    process.exit(1);
}

// Инициализируем S3-клиент для Яндекс Облака
// (протокол S3, но endpoint Яндекса)
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

// Тестовый GET
app.get('/', (req, res) => {
    res.send('Hello World! Спробуйте POST /receive-data');
});

/**
 * POST-запрос:
 * 1) Получаем JSON-данные
 * 2) Открываем Excel-шаблон и записываем данные
 * 3) Генерируем итоговый Excel-файл в буфере
 * 4) Загружаем файл в Яндекс S3
 * 5) Получаем presigned URL для скачивания
 * 6) Возвращаем клиенту JSON с presigned URL
 */
app.post('/receive-data', async (req, res) => {
    console.log('Получен POST-запрос на /receive-data');
    try {
        // Данные, присланные из Airtable (или другого клиента)
        const data = req.body;
        console.log('Получены данные:', data);

        // Открываем Excel-шаблон
        const workbook = await XlsxPopulate.fromFileAsync(excelFilePath);
        const sheet = workbook.sheet(0);

        // Пример записи данных (B10, D10, E10, J10). Настройте под себя:
        sheet.cell('B10').value(data.Дата || '');
        sheet.cell('D10').value(data.Контрагент?.[0]?.name || '');
        sheet.cell('E10').value(data.Получатель?.[0]?.name || '');
        sheet.cell('J10').value(data["Сумма USD"] || 0);

        // Генерируем в буфер (не сохраняя на диск)
        const buffer = await workbook.outputAsync();

        // Формируем уникальное имя файла (Key) для загрузки в S3
        const s3Key = `uploads/updatedFile-${Date.now()}.xlsx`;

        // Отправляем файл в бакет Яндекс Облака
        await s3Client.send(new PutObjectCommand({
            Bucket: bucketName,
            Key: s3Key,
            Body: buffer,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }));

        console.log(`Файл загружен: s3://${bucketName}/${s3Key}`);

        // Генерируем ссылку на скачивание (presigned URL), действующую, например, 1 час
        const getObjectParams = {
            Bucket: bucketName,
            Key: s3Key
        };
        const fileUrl = await getSignedUrl(
            s3Client,
            new GetObjectCommand(getObjectParams),
            { expiresIn: 3600 } // время в секундах
        );

        // Возвращаем JSON с ссылкой
        return res.json({
            success: true,
            message: 'File created and uploaded successfully',
            fileUrl
        });

    } catch (error) {
        console.error('Ошибка при обработке данных:', error);
        res.status(500).json({
            success: false,
            message: 'Ошибка при обработке данных'
        });
    }
});

// Пример GET-запроса на скачивание локального template.xlsx (для отладки)
app.get('/download-latest', (req, res) => {
    const file = path.join(__dirname, 'template.xlsx');
    res.download(file, 'Updated.xlsx', (err) => {
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
