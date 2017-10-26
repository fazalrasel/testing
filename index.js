const parseXlsx = require('excel');
const Promise = require('bluebird');
const clipboard = require('copy-paste');
const knex = require('./db/knex');
const robot = require('robotjs');
let Excel = require('exceljs');
const fs = require('fs-extra');
const argv = require('yargs').argv;

robot.setMouseDelay(200); // wait for another 500ms

function toggleTab() {
  robot.keyToggle('alt', 'down');
  robot.keyTap('tab');
  // robot.keyTap('tab');
  robot.keyTap('enter');
  robot.keyTap('alt');
}

const config = {
  fileName: 'target.xlsx',
  worksheetNumber: 1,
  log: false,

  glass: [440, 530],
  glassColor: '81d0e3',

  searchBox: [285, 200],

  greenArrow: [440, 530],
  greenArrowColor: '85bb51',

  topOfList: [475, 180],

  newWindowEdit: [230, 145],
  newWindowCopyPage: [275, 310],

  newWindowClose: [1275, 120],

  newWindowColorCheck: [1160, 150],
  newWindowColor: 'd1d8e3',
};

toggleTab();

robot.moveMouse(110, 185);
robot.mouseToggle("down");
robot.dragMouse(10, 90);
robot.mouseToggle("up");

robot.moveMouse(230, 430);
robot.mouseToggle("down");
robot.dragMouse(230, 555);
robot.mouseToggle("up");

const time = argv.time || 80;
// toggleTab();// robot.keyTap('c', 'control');

parseXlsx(config.fileName, config.worksheetNumber, function (err, data) {
  if (err) throw err;
  // data is an array of arrays
  const filtered = data.filter((d) => d[8] === ''); // notifier name is empty
  let newArray = [];
  filtered.forEach((d) => {
    // find if existed on newArray
    const find = newArray.find((i) => i[4] === d[4]);
    if (find === undefined) {
      newArray.push(d);
    }
  });
  // console.log(newArray);
  Promise.mapSeries(newArray, (container) => {
    // map on every container
    if (container[4] === '') {
      throw Error('Container number is empty');
    }

    /*
     Step: 1
     Check if we are on the search container panel.
     We are taking decision based on the color of the SEARCH Glass At Bottom
     */
    let checkSearchGlass = new Promise((resolve) => {
      let tryoutCounter = 1;
      const searchGlass = setInterval(() => {
        if (config.log) {
          console.log(robot.getPixelColor(config.glass[0], config.glass[1]), config.glass[0], config.glass[1]);
        }
        if (robot.getPixelColor(config.glass[0], config.glass[1]) === config.glassColor) {
          clearInterval(searchGlass);
          resolve(true);
        }
        tryoutCounter += 1;
        if (tryoutCounter > (time * 4)) {
          // after 60 second
          clearInterval(searchGlass);
          resolve(false);
        }
      }, 500);
    });


    /*
     Step: 2
     Check For Green LEFT Arrow which is on the details windows of way bills.

     */
    let checkGreenLeftArrow = new Promise((resolve) => {
      let tryoutCounter = 1;
      const greenLeftArrow = setInterval(() => {
        if (config.log) {
          console.log(robot.getPixelColor(config.greenArrow[0], config.greenArrow[1]), config.greenArrow[0], config.greenArrow[1]);
        }
        if (robot.getPixelColor(config.greenArrow[0], config.greenArrow[1]) === config.greenArrowColor) {
          clearInterval(greenLeftArrow);
          resolve(true);
        }
        tryoutCounter += 1;
        if (tryoutCounter > (time * 5)) {
          // after 30 second
          clearInterval(greenLeftArrow);
          resolve(false);
        }
      }, 500);
    });

    return checkSearchGlass
      .then((readyToSearch) => {
        if (readyToSearch) {
          robot.moveMouse(config.searchBox[0], config.searchBox[1]);
          robot.mouseClick('left', true);
          robot.typeString(container[4].replace(/\s+/g, ''));
          robot.moveMouse(config.glass[0], config.glass[1]);
          robot.mouseClick('left');
          return true;
        } else {
          throw  Error('Not Ready To Search Container');
        }
      })
      .then(() => checkGreenLeftArrow)
      .then((listArrived) => {
        if (listArrived) {
          // copy list and transfer to array;
          robot.moveMouse(config.topOfList[0], config.topOfList[1]);
          robot.mouseClick('left');
          robot.keyTap('a', 'control');
          robot.keyTap('c', 'control');
          robot.moveMouse(config.topOfList[0], config.topOfList[1]);
          robot.mouseClick('left');
          return clipboard.paste().split('\n').map((i) => i.split('\t'));
        } else {
          throw  Error('Taking To Long To Fetch Lists');
        }
      })
      .then((lists) => {
        lists = lists.slice(0, 20);
        return Promise.mapSeries(lists, (list, index) => {
          // console.log(list);
          // look for color which will indicate that new window opened
          let checkForNewWindowOpen = new Promise((resolve) => {
            let tryoutCounter = 1;
            const checkForNewWindow = setInterval(() => {
              if (robot.getPixelColor(config.newWindowColorCheck[0], config.newWindowColorCheck[1]) === config.newWindowColor) {
                clearInterval(checkForNewWindow);
                resolve(true);
              }
              tryoutCounter += 1;
              if (tryoutCounter > (time * 3)) {
                // after 30 second
                clearInterval(checkForNewWindow);
                resolve(false);
              }
            }, 500);
          });

          return checkGreenLeftArrow
            .then((readyForLeftClick) => {
              if (readyForLeftClick) {
                robot.moveMouse(config.topOfList[0], config.topOfList[1] - 2 + ((index + 1) * 15)); // move to single row for right context list
                robot.mouseClick('right'); // right click for context list
                robot.moveMouse(config.topOfList[0] + 8, config.topOfList[1] - 2 + ((index + 1) * 15) + 8); // position on context list
                // robot.setMouseDelay(10000);
                robot.mouseClick('left'); // click on 'view way bill'
                return true;
              } else {
                throw Error('Not Ready to right click and click on view on way bill');
              }
            })
            .then(() => checkForNewWindowOpen)
            .then((newWindowOpened) => {
              if (newWindowOpened) {
                // robot.setMouseDelay(0);
                robot.moveMouse(config.newWindowEdit[0], config.newWindowEdit[1]);
                robot.mouseClick('left');
                robot.moveMouse(config.newWindowCopyPage[0], config.newWindowCopyPage[1]);
                robot.mouseClick('left');

                robot.moveMouse(config.newWindowClose[0], config.newWindowClose[1]);
                robot.mouseClick('left');
                let str = clipboard.paste();
                const findCode = str.match(/COD:(.*)/ig);
                const manifested_package = str.match(/NBR:(.*)/ig); // package size
                const hsb_msb_code = findCode[1].replace('COD: ', '');
                const port_of_origin_code = findCode[2].replace('COD: ', '');
                const voyage = str.match(/VOY:.*/g);
                const voyage_number = voyage[0].replace('VOY: ', '');

                const line = str.match(/LIN:.*/g);
                const line_number = line[0].replace('LIN: ', '');

                const dates = str.match(/DAT:.*/g);
                const findName = str.match(/NAM:(.*)/ig);
                const notifier_name = findName[10].replace('NAM: ', '');
                const a = str.split('\n').map((i) => i.replace(/\r?\n|\r/g, ""));

                const consignor_name = findName[6].replace('NAM: ', '');
                // const newA = a.slice(38, 45);
                const findAddress = a.slice(35, 55)
                  .join(' ')
                  .match(/ADD:.*COD:.\d/);

                // const newA = a.slice(38, 45);
                const consignor_address = a.slice(27, 37)
                  .join(' ')
                  .match(/ADD:.*COD:.\d/);

                const des = str.match(/DSC:(.*)/ig);
                const desToInsert = des[1].replace('DSC: ', '');
                // const notifierAddress = findAddress;
                const obj = {
                  hsb_msb_code,
                  notifier_name,
                  notifier_address: findAddress ? findAddress[0].split('COD')[0].replace('ADD: ', '') : '',
                  consignor_name,
                  consignor_address: consignor_address ? consignor_address[0].split('COD')[0].replace('ADD: ', '') : null,
                  container_number: container[4].replace(/\s+/g, ''),
                  vessel_name: list[5],
                  b_l_ref: list[3],
                  voyage_number,
                  line_number,
                  port_of_origin_code,
                  container_product_description: desToInsert,
                  departure_date: dates[0].replace('DAT: ', ''),
                  arrival_date: dates[1].replace('DAT: ', ''),
                  package_name: findName[12].replace('NAM: ', ''),
                  manifested_package: manifested_package[1].replace('NBR: ', ''),
                };
                // console.log(obj);
                return knex('containers')
                  .where({
                    container_number: container[4].replace(/\s+/g, ''),
                    b_l_ref: list[3],
                  })
                  .first()
                  .then((found) => {
                    if (found) return 'ok';
                    return knex('containers')
                      .insert(Object.assign({}, obj, { raw: str }))
                      .then((inserted) => {
                        console.log(inserted);
                        return inserted;
                      })
                  })
              } else {
                throw Error('Check Network. Fail To Open Individual Notifier Info');
              }
            })
            .then(() => checkGreenLeftArrow);
        })
      })
      .then((rows) => {
        // single container fetching complete
        // press on green botton
        robot.moveMouse(config.greenArrow[0], config.greenArrow[1]);
        robot.mouseClick('left');
        return rows;
      });
  })
    .then(() => {
      return insertToExcel(data);
    })
    .catch((err) => {
      return insertToExcel(data, err);
    })
});

function insertToExcel(data, err) {
  toggleTab();
  if (err) {
    console.log(err);
  }
  return Promise.mapSeries(data, (originalData) => {
    // console.log(originalData);
    // every data is in array format;
    const packageInfo = originalData[12].split(' ');
    const manifestedPackageWhole = packageInfo.pop();
    const manifestedPackage = manifestedPackageWhole ? manifestedPackageWhole.slice(0, -2) : null;
    const packageName = packageInfo.join(' ');
    // console.log(packageInfo, packageName, manifestedPackage, originalData[12]);
    // console.log(originalData[20]);
    // console.log(originalData[20].replace(/\s+/g, ''));
    return knex('containers')
      .where({
        container_number: originalData[4], // container column on excel file
        package_name: packageName,
        manifested_package: manifestedPackage,
        // container_product_description: originalData[20], // description on excel sheet
      })
      .andWhereRaw('replace(container_product_description, " ", "") = ?', [originalData[20].replace(/\s+/g, '')])
      // .debug()
      .first()
      .then((databaseData) => {
        // console.log(databaseData);
        if (databaseData) {
          originalData[6] = databaseData.consignor_name;
          originalData[7] = databaseData.consignor_address;
          originalData[5] = databaseData.port_of_origin_code;
          originalData[8] = databaseData.notifier_name;
          originalData[9] = databaseData.notifier_address;
          return originalData;
        } else {
          return originalData;
        }
      })
  })
    .then((modifiedData) => {
      // get all modified row. save on excel file;
      let workbook = new Excel.Workbook();
      let worksheet = workbook.addWorksheet('1');
      worksheet.addRows(modifiedData);
      return workbook.xlsx.writeFile('t.xlsx');
    })
    .then(() => {
      return fs.move('./t.xlsx', './target.xlsx', { overwrite: true });
    })
    .then(() => {
      return fs.remove('./t.xlsx');
    })
}
