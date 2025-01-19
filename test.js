const fetch = require('node-fetch'); // Убедитесь, что эта библиотека установлена: npm install node-fetch

(async () => {
    const jsonString = `{
        "id": "rec5YII8ZtUN0gTDI",
        "№": 2565,
        "Контрагент": [{"id": "recPDPvhR8dKEMih4", "name": "ТДК"}],
        "Дата": "2025-01-17",
        "Отправитель": [{"id": "recTNbN78V22GCISr", "name": "ТДК"}],
        "Контрагенты": ["recPDPvhR8dKEMih4"],
        "Получатель": [{"id": "recmbuJHoL3kW3rug", "name": "RR LOGISTIC"}],
        "Валюта": {"id": "selXXDDnsxg2UXN8p", "name": "USD", "color": "cyanLight2"},
        "Сумма": -289632.7,
        "Сумма RUB": null,
        "Сумма USD": -289632.7,
        "Сумма USDT": null,
        "Сумма CNY": null,
        "Сумма AED": null,
        "Сумма EURO": null,
        "Ордеры": null
    }`;

    // Преобразование JSON-строки в объект
    const record = JSON.parse(jsonString);

    // 4. Извлекаем данные из необходимых полей
    let fields = {
        id: record.id,
        "№": record["№"],
        Контрагент: record["Контрагент"],
        Дата: record["Дата"],
        Отправитель: record["Отправитель"],
        Контрагенты: record["Контрагенты"],
        Получатель: record["Получатель"],
        Валюта: record["Валюта"],
        Сумма: record["Сумма"],
        "Сумма RUB": record["Сумма RUB"],
        "Сумма USD": record["Сумма USD"],
        "Сумма USDT": record["Сумма USDT"],
        "Сумма CNY": record["Сумма CNY"],
        "Сумма AED": record["Сумма AED"],
        "Сумма EURO": record["Сумма EURO"],
        Ордеры: record["Ордеры"]
    };

    // 5. Логируем значения каждой переменной перед формированием JSON
    console.log("Извлечённые данные:");
    console.log("ID:", fields.id);
    console.log("№:", fields["№"]);
    console.log("Контрагент:", fields.Контрагент);
    console.log("Дата:", fields.Дата);
    console.log("Отправитель:", fields.Отправитель);
    console.log("Контрагенты:", fields.Контрагенты);
    console.log("Получатель:", fields.Получатель);
    console.log("Валюта:", fields.Валюта);
    console.log("Сумма:", fields.Сумма);
    console.log("Сумма RUB:", fields["Сумма RUB"]);
    console.log("Сумма USD:", fields["Сумма USD"]);
    console.log("Сумма USDT:", fields["Сумма USDT"]);
    console.log("Сумма CNY:", fields["Сумма CNY"]);
    console.log("Сумма AED:", fields["Сумма AED"]);
    console.log("Сумма EURO:", fields["Сумма EURO"]);
    console.log("Ордеры:", fields.Ордеры);

    // 6. Формируем JSON-объект
    let jsonData = JSON.stringify(fields);
    console.log("Сформированный JSON:", jsonData);

    // 7. Отправляем JSON через fetch
    let endpoint = "http://localhost:8080/receive-data"; // Укажи свой локальный/удалённый сервер
    let response = await fetch(endpoint, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: jsonData
    });

    // 8. Обрабатываем ответ сервера
    if (response.ok) {
        let responseData = await response.json();
        console.log("Ответ от сервера:", responseData);
    } else {
        console.error("Ошибка при отправке данных:", response.status, response.statusText);
    }
})();
