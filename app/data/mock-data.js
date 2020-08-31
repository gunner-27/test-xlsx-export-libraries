const sheet1 = {
  name: "Красивый лист",
  headers: [
    {
      key: "fio",
      text: "ФИО",
      width: 30,
    },
    {
      key: "date",
      text: "ДАТА",
      width: 10,
    },
    {
      key: "num",
      text: "ЧИСЛО",
      width: 10,
    },
    {
      key: "text",
      text: "ТЕКСТ",
      width: 50,
    },
    {
      key: "status",
      text: "СТАТУС",
      width: 12,
    },
    {
      key: "colored_num",
      text: "ЦВЕТНОЕ ЧИСЛО",
      width: 24,
    },
  ],
  styles: {
    headers: {
      font: {
        name: "Calibri",
        size: 18,
        bold: true,
      },
      alignment: {
        vertical: "middle",
        horizontal: "center",
      },
    },
  },
  data: [
    {
      FIO: "Тестовый Агент Иванович",
      date: "27.12.2019",
      num: 144678,
      text: "Длинное название услуги из какого-то пакета",
      status: "В работе",
      colored_num: 12,
    },
    {
      FIO: "Реферрерский Реферер",
      date: "07.02.2019",
      num: 1730,
      text: "Другое название услуги из какого-то пакета",
      status: "Архивирован",
      colored_num: 948,
    },
    {
      FIO: "Петр Петрович Тестиков",
      date: "01.01.2019",
      num: 730,
      text: "Услуга из пакета (короткая)",
      status: "В работе",
      colored_num: 1234,
    },
  ],
};

const sheet2 = {
  name: "Другой лист",
  headers: [
    {
      key: "fio",
      text: "Имя Герой Фамилия",
      width: 30,
    },
    {
      key: "status",
      text: "СТАТУС",
      width: 12,
    },
    {
      key: "date",
      text: "ДАТА",
      width: 10,
    },
    {
      key: "text",
      text: "ОПИСАНИЕ",
      width: 50,
    },
    {
      key: "sex",
      text: "ПОЛ",
      width: 7,
    },
  ],
  styles: {
    headers: {
      font: {
        name: "Calibri",
        size: 18,
        bold: true,
      },
      headersColors: [
        "FFFF0000",
        "FFec6333",
        "fffdf731",
        "ff7ff21a",
        "ff58ccfb",
      ],
      alignment: {
        vertical: "middle",
        horizontal: "center",
      },
    },
  },
  data: [
    {
      FIO: "Тони ЖелезныйЧеловек Старк",
      status: "Мертв",
      date: "27.12.2019",
      text: "Богатый, стреляет лазерами и ракетами",
      sex: "М",
    },
    {
      FIO: "Питер ЧеловекПаук Паркер",
      status: "Жив",
      date: "07.02.2019",
      text: "Смешной с паутиной",
      sex: "М",
    },
    {
      FIO: "Брюс Халк Бенер",
      status: "Жив",
      date: "01.01.2019",
      text: "Большой и зеленый",
      sex: "-",
    },
    {
      FIO: "Наташа ЧернаяВдова Романова",
      status: "Жив",
      date: "02.01.2019",
      text: "Смертоностный шпион, владеет боевыми искусствами и оружием",
      sex: "Ж",
    },
  ],
};

module.exports = { sheet1, sheet2 };
