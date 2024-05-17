// GENERAL CONSTANTS
const workbook = SpreadsheetApp.getActiveSpreadsheet();
const allSheets = workbook.getSheets();
const ui = SpreadsheetApp.getUi();

const months = {
  "январь": 1,
  "февраль": 2,
  "март": 3,
  "апрель": 4,
  "май": 5,
  "июнь": 6,
  "июль": 7,
  "август": 8,
  "сентябрь": 9,
  "октябрь": 10,
  "ноябрь": 11,
  "декабрь": 12,
};
///


// COMMON FUNCRIONS

// возвращает список листов с по или false если какого-то не хватает
function getCorrectNames(startMonth, startYear, endMonth, endYear, check="РО") {
  const toMonth = months[endMonth];
  const fromMonth = months[startMonth];
  const totalStartMonth = startYear * 12 + fromMonth
  const totalEndMonth = endYear * 12 + toMonth
  if (totalEndMonth < totalStartMonth) {
    ui.alert("Дата окончания не может быть раньше дата начала")
    return false
  }

  function isListFits(list) {
    const name = list.getName()
    if (check === "РО" && name.slice(0, 2) !== check) return false
    if (check === "Щебень" && name.slice(0, 6) !== check) return false
    const listYear = check === "РО" ?  Number(name.slice(3, 7)) : Number(name.slice(14, 18))
    const listMonth = check === "РО" ? Number(name.slice(8, 10)) : Number(name.slice(19, 21))
    const totalListMonth = listYear * 12 + listMonth
    return totalListMonth >= totalStartMonth && totalListMonth <= totalEndMonth
  }

  const correctLists = allSheets.filter(isListFits)
  const numOfMonths = totalEndMonth - totalStartMonth + 1
  const absentSheetNum = numOfMonths - correctLists.length
  if (absentSheetNum > 0) {
    ui.alert(`Не хватает ${absentSheetNum} листов импортируйте данные или измените диапазон`)
    return false
  }
  if(absentSheetNum < 0) {
    ui.alert("Есть листы за одну и ту же дату, удалите или переименуйте неактуальный")
    return false
  }
  if(!correctLists.length) return false
  return correctLists.map(list => list.getName())
}

function getRightEnding(num_s) {
  let ending = "";
  const i_end = [2, 3, 4];
  const ok_end = [0, 5, 6, 7, 8, 9];
  if ((num_s > 10) & (num_s < 20)) {
    ending = "отгрузок";
  } else {
    if (num_s % 10 === 1) {
      ending = "отгрузка";
    } else if (i_end.includes(num_s % 10)) {
      ending = "отгрузки";
    } else if (ok_end.includes(num_s % 10)) {
      ending = "отгрузок";
    }
  }
  return ending;
}
////

// BUYER REPORT
function buyerReport() {
  const data = {};
  let totalAmount = 0;
  let totalCost = 0;
  let totalPrice = 0;
  
  const reportSheet = workbook.getSheetByName("Покупатели Отчет");
  const nameChosen = reportSheet.getRange("B2").getCell(1, 1).getValue();
  const fractionChosen = reportSheet.getRange("B3").getCell(1, 1).getValue();
  const startMonth = reportSheet.getRange("C4").getCell(1, 1).getValue();
  const endMonth = reportSheet.getRange("C5").getCell(1, 1).getValue();
  const startYear = Number(reportSheet.getRange("D4").getCell(1, 1).getValue());
  const endYear = Number(reportSheet.getRange("D5").getCell(1, 1).getValue());

  function getDataFromSheet(sheetName) {
    const ss = workbook.getSheetByName(sheetName);
    const range = ss.getRange("A5:AK");
    const before2024 = Number(sheetName.slice(3, 7)) === 2023

    function filterByName(row) {
      const name = before2024 ? row[17] : row[9]
      if (nameChosen === "ВСЕ") return name === "" ? false : true;
      if (name === nameChosen) return true;
      return false;
    }

    function filterByFraction(row) {
      if (fractionChosen === "ВСЕ") return true;
      const fraction = before2024 ? row[18].trim() : row[10].trim()
      return (fraction === fractionChosen) 
    }

    const listdata = range
      .getValues()
      .filter(filterByName)
      .filter(filterByFraction);

    listdata.forEach(row => {
      const buyer = before2024 ? row[17] : row[9];
      const fraction = before2024 ? row[18].trim() : row[10].trim();
      const amount = before2024 ? row[19] : row[11];
      const p = before2024 ? row[26] : row[17]
      const gross_price = p * amount;
      const pp = before2024 ? row[33] : row[23];
      const net_price = pp * amount;
      const station = before2024 ? row[14] : row[7];

      if (!(fraction in data)) {
        data[fraction] = {};
      }

      if (!(station in data[fraction])) {
        data[fraction][station] = {
          total_money_gross: gross_price,
          total_money_net: net_price,
          total_amount: amount,
          average_net: 0,
          average_gross: 0,
          buyers: {},
        };
      }

      if (!(buyer in data[fraction][station].buyers)) {
        data[fraction][station].buyers[buyer] = {
          amount: 0,
          gross: 0,
          net: 0,
          shipments: [],
        };
      }

      data[fraction][station].buyers[buyer].amount += amount;
      data[fraction][station].buyers[buyer].gross += gross_price;
      data[fraction][station].buyers[buyer].net += net_price;
      data[fraction][station].buyers[buyer].shipments.push(row);
      if ((buyer !== "ТДГЖ") & (net_price > 0)) {
        data[fraction][station].total_money_gross += gross_price;
        data[fraction][station].total_money_net += net_price;
        data[fraction][station].total_amount += amount;
      }
      totalAmount += amount;
      totalCost += gross_price;
      totalPrice += net_price;

    })
  } 
  // конец функции getData

  const listOfMonths = getCorrectNames(startMonth, startYear, endMonth, endYear);
  
  if (!listOfMonths)  return

  listOfMonths.forEach(getDataFromSheet)

  // Вычисляем средние по станции перед репортом
  for (const fraction in data) {
    data[fraction]["fraction_total_amount"] = 0;
    data[fraction]["fraction_total_money_net"] = 0;
    for (const station in data[fraction]) {
      data[fraction][station].average_net = data[fraction][station].total_money_net / data[fraction][station].total_amount;
      data[fraction][station].average_gross = data[fraction][station].total_money_gross / data[fraction][station].total_amount;
      if (
        station !== "fraction_total_amount" &&
        station !== "fraction_total_money_net"
        ) {
          data[fraction]["fraction_total_amount"] += data[fraction][station].total_amount;
          data[fraction]["fraction_total_money_net"] += data[fraction][station].total_money_net;
        }
    }
  }

  displayBuyersReport();


  function displayBuyersReport() {
    reportSheet.deleteRows(9, 1992);
    reportSheet.insertRowsAfter(8, 1992);

    let i = 1; // в range нумерация с 1
    for (fraction in data) {
      const av_price =
        data[fraction].fraction_total_money_net /
        data[fraction].fraction_total_amount;
      for (station in data[fraction]) {
        for (buyer in data[fraction][station].buyers) {
          let groupStart = i + 8;
          let groupEnd =
            groupStart +
            data[fraction][station].buyers[buyer].shipments.length -
            1;
          let k = 0;
          do {
            // записали все отгрузки по одной
            const shipment = data[fraction][station].buyers[buyer].shipments[k];
            const rangeToSet = reportSheet.getRange(`A${i + 8}:K${i+8}`)

            const before2024 = shipment[0] < new Date('01.01.2024')

            const res = before2024 ?
            [[
              shipment[0], shipment[18], shipment[14], 
              shipment[17], shipment[19], shipment[26],
              shipment[26] * shipment[19], shipment[33],
              data[fraction][station]["average_net"], av_price,
              `=100%*(I${i + 8}-J${i + 8})/J${i + 8}`
            ]] 
            :
            [[
              shipment[0], shipment[10] , shipment[7],
              shipment[9], shipment[11], shipment[17],
              shipment[11] * shipment[17], shipment[23],
              data[fraction][station]["average_net"], av_price,
              `=100%*(I${i + 8}-J${i + 8})/J${i + 8}`
            ]]
            rangeToSet.setValues(res);
            //

            i++;
            k++;
          } while (i < groupEnd - 7);
          // свернули все отгрузки 
          let rangeToGroup = reportSheet.getRange(
            groupStart,
            1,
            groupEnd - groupStart + 1
          );
          rangeToGroup.shiftRowGroupDepth(1);
          let group = reportSheet.getRowGroup(groupStart, 1);
          group.collapse();
          //

          // записали результирующее
          let num_s =
            data[fraction][station]["buyers"][buyer]["shipments"].length;
          let ending = getRightEnding(num_s);
          const shp = `${num_s} ${ending}`;
          const net = data[fraction][station].buyers[buyer].net;
          const amount = data[fraction][station].buyers[buyer].amount;
          const gross = data[fraction][station].buyers[buyer].gross;

          const rangeToSet = reportSheet.getRange(`A${i + 8}:K${i + 8}`);
          
          const res = [[shp, fraction, station, buyer, amount, gross/amount, gross, net/amount, data[fraction][station].average_net, av_price, `=100%*(I${i + 8}-J${i + 8})/ABS(J${i + 8})`]];
          rangeToSet.setValues(res);

          // 

          i++;
        }
      }
    }
    // дописали сумму в шапку таблицы
    reportSheet.getRange("E8:E8").getCell(1, 1).setValue(totalAmount);
    reportSheet.getRange("G8:G8").getCell(1, 1).setValue(totalCost);

    reportSheet.autoResizeRows(9, i - 1);
    ui.alert(`Отчет сформирован`)
  }
}

/// BUYERS REPORT END


function carrierReport() {
  const data = {};
  const station_price = {};
  const reportSheet = workbook.getSheetByName("Перевозчики Отчет");
  const nameChosen = reportSheet.getRange("B2").getCell(1, 1).getValue();
  const stationChosen = reportSheet.getRange("B3").getCell(1, 1).getValue();
  const startMonth = reportSheet.getRange("C4").getCell(1, 1).getValue();
  const endMonth = reportSheet.getRange("C5").getCell(1, 1).getValue();
  const startYear = Number(reportSheet.getRange("D4").getCell(1, 1).getValue());
  const endYear = Number(reportSheet.getRange("D5").getCell(1, 1).getValue());

  const listOfMonths = getCorrectNames(startMonth, startYear, endMonth, endYear);

  function getDataFromSheet(sheetName) {
    const ss = workbook.getSheetByName(sheetName);
    const range = ss.getRange("A5:AK");
    const before2024 = Number(sheetName.slice(3, 7)) === 2023

    function filterByName(row) {
      const name = before2024 ? row[17] : row[9]
      if (nameChosen === "ВСЕ") return name === "" ? false : true;
      if (name === nameChosen) return true;
      return false;
    }

    function filterByStation(row) {
      if (stationChosen === "ВСЕ") return true
      const station = before2024 ? row[14] : row[7]
      if (station === stationChosen) return true;
      return false;
    }
    
    const listedData = range
      .getValues()
      .filter(filterByName)
      .filter(filterByStation);

    if (!listedData.length) return ui.alert('Не было доставок с заданными параметрами')

    for (row of listedData) {
      const isShip =  before2024 ? !!row[10] : row[4] === 'т/х'


      const carrier =  before2024 ? row[10] + row[7] : row[5];
      const station = before2024 ? row[14] : row[7];
      const carries =  before2024 ? row[4] : row[3];
      const amount = before2024 ? row[19] : row[11];
      const type = before2024 ? row[5] : row[4];
      const deliveryPriceNet = before2024 ? Number(row[22]) + Number(row[23]) : Number(row[13]);
      const deliveryPriceGross = before2024 ? Number(row[21]) + deliveryPriceNet : Number(row[12]) + deliveryPriceNet;

      if (!(carrier in data)) {
        data[carrier] = {
          stations: {},
          carries: 0,
          amount: 0,
          isShip,
          total_money: 0,
        };
      }

      if (!(station in data[carrier].stations)) {
        data[carrier].stations[station] = {
          station_carries: 0,
          station_amount: 0,
          deliveries: [],
          station_total_price_net: 0,
          station_total_price_gross: 0,
          carrier_type: "",
        };
      }

      data[carrier].carries += carries;
      data[carrier].amount += amount;
      data[carrier].total_money += deliveryPriceNet;
      data[carrier].stations[station].station_carries += carries;
      data[carrier].stations[station].station_amount += amount;
      const cur_type = data[carrier].stations[station].carrier_type;
      if (cur_type === "") {
        data[carrier].stations[station].carrier_type = type;
      } else if (cur_type !== type) {
        data[carrier].stations[station].carrier_type = "разные";
      }
      data[carrier].stations[station].deliveries.push(row);
      data[carrier].stations[station].station_total_price_net += deliveryPriceNet;
      data[carrier].stations[station].station_total_price_gross += deliveryPriceGross;

      if (!(station in station_price)) {
        station_price[station] = { amount: 0, price_net: 0, price_gross: 0 };
      }
      station_price[station].amount += amount;
      station_price[station].price_net += deliveryPriceNet;
      station_price[station].price_gross += deliveryPriceGross;
    }
  }

  function displayResult() {
    reportSheet.deleteRows(9, 992);
    reportSheet.insertRowsAfter(8, 992);

    let i = 1;
    for (carrier in data) {
      for (station in data[carrier].stations) {
        // конкретные отгрузки
        const deliveries = data[carrier].stations[station].deliveries;
       
        const groupStart = i + 8;

        const setRow = (row) => {
          const before2024 = row[0] < new Date('01.01.2024')
          const type = data[carrier].isShip ? 'водой'
            : before2024 
              ? 
              row[5] 
              : 
              row[4]
          const carriers = before2024 ? row[4] : row[3];
          const amount = before2024 ?  row[19] : row[11];
          const carrier_station_price_net =
            data[carrier].stations[station].station_total_price_net / 
            data[carrier].stations[station].station_amount;

          const deliveryPriceNet = before2024 ? Number(row[22]) + Number(row[23]) : Number(row[13]);

          const stationAverageGross = 
            data[carrier].stations[station].station_total_price_gross /
            data[carrier].stations[station].station_amount

          const stationAverageNet = station_price[station].price_net / station_price[station].amount;

          const diff = stationAverageNet ? `=100%*(G${i + 8}-J${i + 8})/ABS(J${i + 8})` : ''

          const res = [[
            row[0], carrier, station, 
            type, carriers, amount, 
            carrier_station_price_net, deliveryPriceNet, stationAverageGross,
            stationAverageNet, diff
          ]]

          const rangeToSet = reportSheet.getRange(`A${i+ 8}:K${i+8}`)
          rangeToSet.setValues(res)

          i++;
        }
        deliveries.forEach(row => setRow(row));

        const groupEnd = i + 7;
        let rangeToGroup = reportSheet.getRange(
          groupStart,
          1,
          groupEnd - groupStart + 1
        );
        rangeToGroup.shiftRowGroupDepth(1);
        let group = reportSheet.getRowGroup(groupStart, 1);
        group.collapse();
        // конкретные отгрузки закончились

        let num = deliveries.length;
        let ending = getRightEnding(num);
        const shp = `${num} ${ending}`;
        const type = data[carrier].isShip ? 'водой' : data[carrier].stations[station].carrier_type;
        const station_carries = data[carrier].stations[station].station_carries;
        const station_amount = data[carrier].stations[station].station_amount;
        const carrier_station_price_net =
            data[carrier].stations[station].station_total_price_net /
            data[carrier].stations[station].station_amount;
        const station_total_price_net = data[carrier].stations[station].station_total_price_net

        const station_price_net =
          station_price[station].price_net / station_price[station].amount;

        const diff = station_price_net ? `=100%*(G${i + 8}-J${i + 8})/ABS(J${i + 8})` : ''

        const res = [[
          shp, carrier, station, type,
          station_carries, station_amount, carrier_station_price_net,
          station_total_price_net, station_amount, station_price_net, diff
        ]]

        const rangeToSet = reportSheet.getRange(`A${i+ 8}:K${i+8}`)
        rangeToSet.setValues(res)

        i++;
      }
    }
    reportSheet.autoResizeRows(9, i - 1);
  }

 

  if (listOfMonths) {
    for (name of listOfMonths) {
      getDataFromSheet(name);
    }
    displayResult();
  }
}

function averagesReport() {
  const data = {};
  const fraction_average = {};
  const reportSheet = workbook.getSheetByName("Ср.цены Отчет");
  const buyerChosen = reportSheet.getRange("B2").getCell(1, 1).getValue();
  const fractionChosen = reportSheet.getRange("B3").getCell(1, 1).getValue();
  const startMonth = reportSheet.getRange("C4").getCell(1, 1).getValue();
  const endMonth = reportSheet.getRange("C5").getCell(1, 1).getValue();
  const startYear = Number(reportSheet.getRange("D4").getCell(1, 1).getValue());
  const endYear = Number(reportSheet.getRange("D5").getCell(1, 1).getValue());

  const listOfMonths = getCorrectNames(startMonth, startYear, endMonth, endYear);

  function getDataFromSheet(sheetName) {
    const ss = workbook.getSheetByName(sheetName);
    const range = ss.getRange("A5:AK");
    const before2024 = Number(sheetName.slice(3, 7)) === 2023

    function filterByName(row) {
      const name = before2024 ? row[17] : row[9];
      if (buyerChosen === "ВСЕ") return name === "" ? false : true;
      if (name === buyerChosen) return true;
      return false;
    }

    function filterByFraction(row) {
      if (fractionChosen === "ВСЕ") return true;
      const fraction = before2024 ? row[18].trim() : row[10].trim();
      return (fraction === fractionChosen) 
    }

    const listData = range
      .getValues()
      .filter(filterByName)
      .filter(filterByFraction);
  
    for (row of listData) {
      const before2024 = row[0] < new Date('01.01.2024')
      const buyer = before2024 ? row[17] : row[9]
      const fraction = before2024 ? row[18].trim() : row[10].trim();
      const amount = before2024 ? row[19] : row[11];
      const netPrice = before2024 ?  Number(row[33]) : Number(row[23])
      const netCost = netPrice * amount;
    
      if (!(buyer in data)) {
        data[buyer] = { fractions: {} };
      }
      if (!(fraction in data[buyer].fractions)) {
        data[buyer].fractions[fraction] = {
          total_amount: 0,
          total_cost_net: 0,
          shipments: [],
        };
      }
      data[buyer].fractions[fraction].total_amount += amount;
      data[buyer].fractions[fraction].total_cost_net += netCost;
      data[buyer].fractions[fraction].shipments.push(row);

      if (!(fraction in fraction_average)) {
        fraction_average[fraction] = { total_amount: 0, total_cost: 0 };
      }
      if ((buyer !== "ТДГЖ") & (netCost > 0)) {
        fraction_average[fraction].total_amount += amount;
        fraction_average[fraction].total_cost += netCost;
      }
    } // конец цикла
  } // конец функции getData

  function displayResult() {
    reportSheet.deleteRows(9, 1992);
    reportSheet.insertRowsAfter(8, 1992);

    let i = 1;
    for (buyer in data) {
      for (fraction in data[buyer].fractions) {
        const curFraction = data[buyer].fractions[fraction]
        const shipments = curFraction.shipments;
        const groupStart = i + 8;
        // конкретные отгрузки
        for (row of shipments) {
          const before2024 = row[0] < new Date('01.01.2024')
          const amount = before2024 ? row[19] : row[11];
          const netPrice = before2024 ?  Number(row[33]) : Number(row[23])
          const fractionAveragePrice = fraction_average[fraction].total_cost / fraction_average[fraction].total_amount
          const diff = `=100%*(E${i + 8}-F${i + 8})/ABS(F${i + 8})`
          const res = [[
            row[0], buyer, fraction, amount, netPrice, fractionAveragePrice, diff
          ]] 

          const rangeToSet = reportSheet.getRange(`A${i + 8}:G${ i+ 8}`);
          rangeToSet.setValues(res)
          i++;
        }

        const groupEnd = i + 7;
        let rangeToGroup = reportSheet.getRange(
          groupStart,
          1,
          groupEnd - groupStart + 1
        );
        rangeToGroup.shiftRowGroupDepth(1);
        let group = reportSheet.getRowGroup(groupStart, 1);
        group.collapse();
        // конкретные отгрузки закончились

        const ending = getRightEnding(shipments.length);
        const shipmentString = `${shipments.length} ${ending}`;
        const byerAveragePriceNet = curFraction.total_cost_net / curFraction.total_amount;
        const averagePriceNet = fraction_average[fraction].total_cost / fraction_average[fraction].total_amount;
        const diff = `=100%*(E${i + 8}-F${i + 8})/ABS(F${i + 8})`
        const res = [[
          shipmentString, buyer, fraction, curFraction.total_amount, 
          byerAveragePriceNet, averagePriceNet, diff
        ]]

        const rangeToSet = reportSheet.getRange(`A${i + 8}:G${ i+ 8}`);
        rangeToSet.setValues(res)
        i++;
      }
    }
    reportSheet.autoResizeRows(9, i - 1);
  } // конец displayResult

  if (listOfMonths) {
    for (name of listOfMonths) {
      getDataFromSheet(name);
    }
    displayResult();
  }

  
} // конец функции отчета

function averageShipmentsReport() {
  const reportSheet = workbook.getSheetByName("Ср.цены по отгрузке Отчет");
  const startMonth = reportSheet.getRange("C3").getCell(1, 1).getValue();
  const endMonth = reportSheet.getRange("C4").getCell(1, 1).getValue();
  const startYear = Number(reportSheet.getRange("D3").getCell(1, 1).getValue());
  const endYear = Number(reportSheet.getRange("D4").getCell(1, 1).getValue());
  const fractionChosen = reportSheet.getRange("B2").getCell(1, 1).getValue().trim();
  const data = {};

  const listOfMonths = getCorrectNames(startMonth, startYear, endMonth, endYear);

  console.log('listOgMonths ->',listOfMonths);

  if (listOfMonths) {
    for (listName of listOfMonths) {
      getDataFromSheet(listName);
    }
    displayResult();
  }

  function getDataFromSheet(sheetName) {
    const before2024 = Number(sheetName.slice(3, 7)) === 2023

    function filterByFraction(row) {
      const fraction = before2024 ? row[18].trim() : row[10].trim();
      if (fractionChosen === "ВСЕ") return !!fraction
      return (fraction === fractionChosen) 
    }

    const ss = workbook.getSheetByName(sheetName);
    const range = ss.getRange("A5:AK");
    const listdata = range.getValues().filter(filterByFraction);

    for (row of listdata) {
      const fraction = before2024 ? row[18].trim() : row[10].trim();
      const price_net = before2024 ? Number(row[33]) : Number(row[23]);
      const amount = before2024 ? Number(row[19]) : Number(row[11]);
      const buyer = before2024 ? row[17] : row[9]
      if ((buyer === "ТДГЖ") || !(Number(price_net > 0))) continue
      
      const shipmentTypeDict = {
        'тх' : 'water',
        'т/х' : 'water',
        'самовывоз' : 'selfshipment',
        "самовывоз розница" : 'selfshipment_retail',
        default : 'carrier'
      }
      const shipmentText = before2024 ? row[5] : row[4]
      const shipment_type = shipmentTypeDict[shipmentText.trim()] || shipmentTypeDict.default;

      const defaultFractionValue = {
        amount: 0, 
        cost: 0, 
        maximum: -1, 
        minimum: 10000000000 
      }
      if(amount === 0) continue

      if (!(fraction in data)) {
        data[fraction] = {
          water: {...defaultFractionValue},
          selfshipment: {...defaultFractionValue},
          carrier: { ...defaultFractionValue },
          selfshipment_retail: {...defaultFractionValue}
        };
      }
      

      data[fraction][shipment_type].amount += amount;
      data[fraction][shipment_type].cost += amount * price_net;
      if (price_net > data[fraction][shipment_type].maximum) data[fraction][shipment_type].maximum = price_net;
      if (price_net < data[fraction][shipment_type].minimum) data[fraction][shipment_type].minimum = price_net; 
    } 
  }

  function displayResult() {
    reportSheet.deleteRows(6, 100);
    reportSheet.insertRowsAfter(5, 100);
    function sorting(a, b) {
      return (
        Number(a[0].slice(0, a[0].indexOf("*")).replace(",", ".")) -
        Number(b[0].slice(0, b[0].indexOf("*")).replace(",", "."))
      );
    }
    const firstHat = [[
      '', 'Вагоны + Вода', '', '', '', 'Вагоны', '', '', '',	'Вода', '', ''	
    ]]
    const secondHat = [[
      '', 'Самовывоз', '', '', '', 'Самовывоз', '', '', '',	'Самовывоз розница', '', ''	
    ]]
    const hat = [[
      'Фракция',	'Тоннаж',	'Цена щебня на тонну', 'Стоимость', '',	'Тоннаж',	'Цена щебня на тонну',	'Стоимость', ''	,	'Тоннаж',	'Цена щебня на тонну','Стоимость'
    ]]
    const sorted_data = Object.fromEntries(Object.entries(data).sort(sorting));
    let i = 6;
    let range = reportSheet.getRange(`A${i}:L${i}`);

    const setHat = (hatToSet) => {
      range = reportSheet.getRange(`A${i}:L${i}`);
      range.setValues(hatToSet);
      range.setHorizontalAlignment("center")
      range = reportSheet.getRange(`B${i}:D${i}`);
      range.setBackground('#d2f1da');
      range.merge();
      range.setFontWeight("bold")
      range = reportSheet.getRange(`F${i}:H${i}`);
      range.merge();
      range = reportSheet.getRange(`J${i}:L${i}`);
      range.merge();
      i++;
      range = reportSheet.getRange(`A${i}:L${i}`);
      range.setValues(hat);
      range.setHorizontalAlignment("center");
      range.setFontWeight("bold");
      range.setBorder(null, null, true, null, null, null)
    }

    setHat(firstHat);
    i++;
    
    for (fraction in sorted_data) {
      const carrierCost = data[fraction].carrier.cost;
      const waterCost = data[fraction].water.cost;
      const carrierAmount = data[fraction].carrier.amount;
      const waterAmount = data[fraction].water.amount;
      const totalAmount  = carrierAmount + waterAmount;
      const totalCost = carrierCost + waterCost;
      const average = totalAmount ? totalCost / totalAmount : '';
      const waterAverage = waterAmount ? waterCost / waterAmount : ''
      const carrierAverage = carrierAmount ? carrierCost / carrierAmount : ''

      const res = [[
        fraction, totalAmount || '', average || '', totalCost || '', '',
        carrierAmount || '', carrierAverage || '', carrierCost || '', '',
        waterAmount || '', waterAverage || '', waterCost || ''
      ]]
      range = reportSheet.getRange(`A${i}:L${i}`)
      range.setValues(res);
      i++
    }
    range = reportSheet.getRange(`E${6}:E${i - 1}`)
    range.setBorder(null, true, null, true, null, null)
    range = reportSheet.getRange(`I${6}:I${i - 1}`)
    range.setBorder(null, true, null, true, null, null)
    
    range = reportSheet.getRange(`A${i}:L${i}`)
    range.setBorder(true, null, true, null, null, null)
    i++;
    const secondPartStart = i;
    setHat(secondHat);
    i++;

    for (fraction in sorted_data) {
      const selfshipmentCost = data[fraction].selfshipment.cost;
      const selfshipment_retailCost = data[fraction].selfshipment_retail.cost;
      const selfshipmentAmount = data[fraction].selfshipment.amount;
      const selfshipment_retailAmount = data[fraction].selfshipment_retail.amount;
      const totalAmount  = selfshipmentAmount + selfshipment_retailAmount;
      const totalCost = selfshipmentCost + selfshipment_retailCost;
      const average = totalAmount ? totalCost / totalAmount : '';
      const selfshipment_retailAverage = selfshipment_retailAmount ?  selfshipment_retailCost / selfshipment_retailAmount  : ''
      const selfshipmentAverage = selfshipmentAmount ?  selfshipmentCost / selfshipmentAmount  : ''

      const res = [[
        fraction, totalAmount || '', average || '', totalCost || '' , '',
        selfshipmentAmount || '', selfshipmentAverage || '', selfshipmentCost || '', '',
        selfshipment_retailAmount || '', selfshipment_retailAverage || '', selfshipment_retailCost || ''
      ]]
      range = reportSheet.getRange(`A${i}:L${i}`)
      range.setValues(res);
      i++
    }

    range = reportSheet.getRange(`E${secondPartStart}:E${i - 1}`)
    range.setBorder(null, true, null, true, null, null)
    range = reportSheet.getRange(`I${secondPartStart}:I${i - 1}`)
    range.setBorder(null, true, null, true, null, null)
    range = reportSheet.getRange(`A6:L${i - 1}`)
    range.setBorder(null, null, true, true, null, null)
  }

}

function deleteDuplicates() {
  const names = [];
  for (let i = 0; i < allSheets.length; i++) {
    const current_sheet = allSheets[i];
    const current_name = allSheets[i].getName();
    const len = current_name.length;
    if ((current_name[len - 3] === "(") & (current_name[len - 1] === ")")) {
      const nameSliced = current_name.slice(0, len - 4);
      if (names.includes(nameSliced)) {
        const ss = workbook.getSheetByName(nameSliced);
        workbook.deleteSheet(ss);
        current_sheet.setName(nameSliced);
      } else names.push(current_name);
    } else names.push(current_name);
  }
}

function generateInputs() {
  const buyers = new Set();
  const fractions = new Set();
  const carriers = new Set();
  const stations = new Set();
  for (let i = 0; i < allSheets.length; i++) {
    const current_sheet = allSheets[i];
    const current_name = allSheets[i].getName();
    if (current_name.slice(0, 8) === "РО 2023.") {
      const buyers_all = current_sheet.getRange("R5:R").getValues();
      const fractions_all = current_sheet.getRange("S5:S").getValues();
      const carriers_all_carrige = current_sheet.getRange("H5:H").getValues();
      const carriers_all_ship = current_sheet.getRange("K5:K").getValues();
      const stations_all = current_sheet.getRange("O5:O").getValues();
      for (b of buyers_all) {
        if (b[0]) {
          buyers.add(b[0]);
        }
      }
      for (f of fractions_all) {
        if (f[0]) {
          fractions.add(f[0].trim());
        }
      }
      for (c of carriers_all_carrige) {
        if (c[0]) {
          carriers.add(c[0]);
        }
      }
      for (c of carriers_all_ship) {
        if (c[0]) {
          carriers.add(c[0]);
        }
      }
      for (s of stations_all) {
        if (s[0]) {
          stations.add(s[0]);
        }
      }
    }
  }

  const iterator_b = buyers.values();
  const iterator_f = fractions.values();
  const iterator_c = carriers.values();
  const iterator_s = stations.values();
  const inputs = workbook.getSheetByName("Служебный");
  const input_buyers = inputs.getRange("B3:B");
  input_buyers.clear();
  const input_fractions = inputs.getRange("A3:A");
  input_fractions.clear();
  const input_carriers = inputs.getRange("D3:D");
  input_carriers.clear();
  input_stations = inputs.getRange("E3:E");
  for (let i = 1; i < buyers.size + 1; i++) {
    input_buyers.getCell(i, 1).setValue(iterator_b.next().value);
  }
  for (let i = 1; i < fractions.size + 1; i++) {
    input_fractions.getCell(i, 1).setValue(iterator_f.next().value);
  }
  for (let i = 1; i < carriers.size + 1; i++) {
    input_carriers.getCell(i, 1).setValue(iterator_c.next().value);
  }
  for (let i = 1; i < stations.size + 1; i++) {
    input_stations.getCell(i, 1).setValue(iterator_s.next().value);
  }
}

function deliveryReport() {
  const data = {};
  const reportSheet = workbook.getSheetByName("ЩЕБЕНЬ В ПУТИ Отчет");

  const getDataOnTheWay = (sheet) => {
    const onTheWayData = sheet.getRange("A2:AH").getValues();

    for (row of onTheWayData) {
      if (row[1]) {
        const id = row[1].toString().trim();
        const date = row[0];
        const type = row[2];
        const fraction = row[9].toString().trim();
        const amount = Number(row[6]);
        const els = Number(row[14]);
        const deliveryPrice =
          Number(row[15]) + Number(row[16]) + Number(row[17]);
        data[id] = {
          date,
          type,
          fraction,
          amount,
          els,
          deliveryPrice,
          carries: [],
        };
      }
    }
  };

  function getSheetData(sheet) {
    let alert = true;
    const values = sheet.getRange("A5:AK").getValues();
    for (row of values) {
      const id = row[2];
      if (id === "") {
        continue;
      } else if (alert & !(id in data)) {
        // ui.alert(`Внимание лист ${sheet.getName()}, заявка ${id}. Такой заявки нет на листах "ЩЕБЕНЬ В ПУТИ"`);
        // alert = false;
      } else {
        data[id].carries.push(row);
      }
    }
  }

  for (sheet of allSheets) {
    const name = sheet.getName();
    if (name.slice(0, 13) === "Щебень в пути") {
      getDataOnTheWay(sheet);
    }
  }

  for (sheet of allSheets) {
    const name = sheet.getName();
    if (name.slice(0, 8) === "РО 2023.") {
      getSheetData(sheet);
    }
  }

  let i = 1;
  reportSheet.deleteRows(4, 996);
  reportSheet.insertRowsAfter(3, 996);
  let range = reportSheet.getRange("A4:Z");
  // range.clearContent();

  for (let id in data) {
    let groupStart = i;
    for (delivery of data[id].carries) {
      //заявка
      range.getCell(i, 1).setValue(id);
      // дата
      range.getCell(i, 2).setValue(delivery[0]);
      // тип
      range.getCell(i, 3).setValue(delivery[5]);
      // фракция
      range.getCell(i, 4).setValue(delivery[18]);
      // тоннаж
      range.getCell(i, 6).setValue(delivery[19]);
      // елс
      range.getCell(i, 8).setValue(Number(delivery[21]));
      // доставка
      range
        .getCell(i, 10)
        .setValue(
          Number(delivery[22]) + Number(delivery[23]) + Number(delivery[24])
        );
      // покупатель
      range.getCell(i, 11).setValue(delivery[17]);
      i++;
    }
    let groupEnd = i - 1;

    // +3 тут везде так как первые 3 строки служебные
    let rangeToGroup = reportSheet.getRange(
      groupStart + 3,
      1,
      groupEnd - groupStart + 1
    );
    rangeToGroup.shiftRowGroupDepth(1);
    let group = reportSheet.getRowGroup(groupStart + 3, 1);
    group.collapse();

    range.getCell(i, 1).setValue(id);
    range.getCell(i, 2).setValue(data[id].date);
    range.getCell(i, 3).setValue(data[id].type);
    range.getCell(i, 4).setValue(data[id].fraction);
    range.getCell(i, 5).setValue(data[id].amount);
    range
      .getCell(i, 6)
      .setFormula(`=sum(F${groupStart + 3}:F${groupEnd + 3})/E${i + 3}`);
    range.getCell(i, 6).setNumberFormat("0.0%");
    range.getCell(i, 7).setValue(data[id].els);
    range
      .getCell(i, 8)
      .setFormula(`=sum(H${groupStart + 3}:H${groupEnd + 3})/G${i + 3}`);
    range.getCell(i, 8).setNumberFormat("0.0%");
    range.getCell(i, 9).setValue(data[id].deliveryPrice);
    range
      .getCell(i, 10)
      .setFormula(`=sum(J${groupStart + 3}:J${groupEnd + 3})/I${i + 3}`);
    range.getCell(i, 10).setNumberFormat("0.0%");
    i++;
  }

}
