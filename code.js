const { Document,
    AlignmentType, Packer,
    Paragraph,
    Table,
     TableRow,
TableCell, WidthType
  } = docx;

class MyRow{
  constructor(DataMissionCame, Mission){
    this.DataMissionCame = DataMissionCame;
    this.Mission = Mission;
  }
  display(){
    let num = "";
    if (this.DataMissionCame.getMonth() < 9)
      num = "0";
    return this.DataMissionCame.getDate() + "." + num + (this.DataMissionCame.getMonth() + 1) + "." + this.DataMissionCame.getFullYear();
  }
}

// a and b are javascript Date objects
function dateDiffInDays(a, b) {

  const _MS_PER_DAY = 1000 * 60 * 60 * 24;
  // Discard the time and time-zone information.
  const utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

  return Math.floor((utc2 - utc1) / _MS_PER_DAY);
}

function dateCreator(date) {
  return new Date(date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate());
}

function getRandomInt(min = 0, max) {
  return Math.floor(Math.random() * max + min);
}

function mygetDay(date){

  if(date.getDay() === 0){
    return 7;
  }
  return date.getDay();
}

function EveryDayTasks(starttime, endtime, task){
  let start = new Date(starttime)
  let end = new Date(endtime)
  let rows = [];
  while (start <= end)
  {
    let now = dateCreator(start)
    rows.push(new MyRow(now, task));
    start.setDate(start.getDate() + 1);
  }
  return rows;
}

function EveryWeekTasks(starttime, endtime, task){
  let rows = [];
  let start = new Date(starttime)
  let end = new Date(endtime)
  let now = dateCreator(start)

  if (dateDiffInDays(start,end) < 7 && mygetDay(start) < mygetDay(end)){
    now.setDate(now.getDate() + getRandomInt(0, mygetDay(end) - mygetDay(start) + 1))
    rows.push(new MyRow(now, task))
  }
  else{
    //первая неделя
    now.setDate(now.getDate() + getRandomInt(0, 8 - mygetDay(now)))
    rows.push(new MyRow(now, task))
    start.setDate(start.getDate() + 8 - mygetDay(start))

  //последняя неделя
  now = dateCreator(end)
  now.setDate(now.getDate() - getRandomInt(0, mygetDay(now)))
  rows.push(new MyRow(now, task))
  end.setDate(end.getDate() - mygetDay(end))


  while (start <= end)
  {
    //объявление новой now для отсутствия ошибок из-за передачи по ссылке
    now = dateCreator(start)
    now.setDate(now.getDate() + getRandomInt(0,7))
    //добавление в случайный день недели
    rows.push(new MyRow(now, task));
    start.setDate(start.getDate() + 7);
  }
  }
  return rows;
}

function NoTimeTasks(starttime, endtime, task){
  let rows = [];
  let start = new Date(starttime)
  let end = new Date(endtime)
  let my_num = 0;
  let num_of_rows = getRandomInt(1,11);
  let days_between = dateDiffInDays(start, end)

  if (num_of_rows > days_between)
  {
    num_of_rows = getRandomInt(1, days_between + 1);
  }

  let used_nums = [];

  for (let i = 0; i < num_of_rows; i++){
    //отсутствие повторяющихся дат
    do {
      my_num = getRandomInt(1, days_between + 1);
    }while (used_nums.includes(my_num))
    used_nums.push(my_num);

    let now = dateCreator(start)
    now.setDate(now.getDate() + my_num)
    rows.push(new MyRow(now, task))
  }
  return rows;
}

//------------------------------------------ функция startDate.getFullYear() + "-" + (startDate.getMonth() + 1) + "-" + startDate.getDate()
const startDOCX = (e) => {

  let startdate = document.form.startdate.value;
  let enddate = document.form.enddate.value;
  let StartDate = new Date(startdate);
  let EndDate = new Date(enddate);
  if(startdate == "" || enddate == "") {
    alert("Введите данные")
    return 0;
  }
  if(new Date(StartDate) >= new Date(EndDate)) {
    alert("Начальная дата должна быть меньше даты конца периода обслуживания")
    return 0;
  }

//моя таблица
  const table = new Table({
    margin: {
      top: 500,
      bottom: 500,
    },
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({
            text:"№ п / п",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Дата поступления задания",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Способ получения задания",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Содержание задачи",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Сроки оказания услуг",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Статус задачи",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Примечания",
            style: "normalPara"})],
        })
      ],
    }),
    new TableRow({
      children: [
        new TableCell({
          children: [new Paragraph({
            text: "",
            style: "normalPara"
          })],
        }),
        new TableCell({
          children: [new Paragraph({
            text: "2",
            style: "normalPara"
          })],
        }),
        new TableCell({
          children: [new Paragraph({
            text: "3",
            style: "normalPara"
          })],
        }),
        new TableCell({
          children: [new Paragraph({
            text: "4",
            style: "normalPara"
          })],
        }),
        new TableCell({
          children: [new Paragraph({
            text: "5",
            style: "normalPara"
          })],
        }),
        new TableCell({
          children: [new Paragraph({
            text: "6",
            style: "normalPara"
          })],
        }),
        new TableCell({
          children: [new Paragraph({
            text: "7",
            style: "normalPara"
          })],
        })
      ]
    })
  ],
});

  let rows = EveryDayTasks(startdate, enddate, "Ежедневная проверка актуальности расписания в разделе ИНФОКИОСК");
  rows.push(...EveryDayTasks(startdate, enddate, "Ежедневная проверка актуальности расписания на сайте учреждения"))
  rows.push(...EveryDayTasks(startdate, enddate, "Ежедневная проверка актуальности расписания на электронных табло"))
  rows.push(...EveryWeekTasks(startdate, enddate, "Еженедельная верификация баз данных"))
  rows.push(...EveryWeekTasks(startdate, enddate, "Еженедельная проверка актуальности законодательной базы и актуализация информации"))
  rows.push(...NoTimeTasks(startdate, enddate, "Добавление новостей организации"))
  rows.push(...NoTimeTasks(startdate, enddate, "Размещение визуальных и текстовых материалов по заданию Заказчика"))
  rows.push(...NoTimeTasks(startdate, enddate, "Выполнение регламентных работ по обслуживанию серверов"))

  for (let i = 0; i < rows.length; i++){
    table.root.push(
        new TableRow({
        children: [
        new TableCell({
          children: [new Paragraph({
            text:(i + 1).toString(),
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:rows[i].display(),
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Электронный",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text: rows[i].Mission,
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"В соответствие с договором",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Выполнена",
            style: "normalPara"})],
        }),
        new TableCell({
          children: [new Paragraph({
            text:"Отсутствует",
            style: "normalPara"})],
        })
      ],
    })
    )
  }

const doc = new Document({
  styles: {
    paragraphStyles: [
      {
        id: "normalPara",
        name: "Normal Para",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: "Times New Roman",
          size: 24
        },
        paragraph: {
          alignment: AlignmentType.CENTER,
        },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: 1122,
            right: 1122,
            bottom: 1122,
            left: 1122,
              footer:720,
              header:720
          },
            size: {
              width: 16817, //29660 / 52,31 = 5,67
              height: 11901, // 20990 / 37,02 = 5,67
            }
        },
      },
      children: [
        new Paragraph({
          text: "Отчет об оказанных услугах",
          style: "normalPara",
        }),
        new Paragraph({
          text: "по договору №____________ от____________________",
          style: "normalPara",
        }),
        new Paragraph({
          text: "за период с " + startdate + " по " + enddate,
          style: "normalPara",
        }),
        new Paragraph({
          text: "В соответствии с Договором №________от________________ в отчетном периоде с " + startdate + " по " + enddate + " Исполнитель оказал Заказчику следующие услуги: \n",
          style: "normalPara",
        }),
          table,
        new Paragraph({
          text: "Сдал Исполнитель:                                                                                    Принял Заказчик:",
          style: "normalPara",
        }),
        new Paragraph({
          text: "«_____»___________20_________г.                                                         «_______»_______________20________г.",
          style: "normalPara",
        }),
        new Paragraph({
          text: "М.П.                                                                                                              М.П.",
          style: "normalPara",
        })
      ],
    },
  ]
})

    Packer.toBlob(doc).then(blob =>{
        saveAs(blob, '1.docx')
    })
}

let button = document.form.button;
button.addEventListener("click", startDOCX)

