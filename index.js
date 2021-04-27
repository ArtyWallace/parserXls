const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
  sourceFile: __dirname + '/23.04.21 Челентано.xlsx'
});

let simpleCredit = 0;
let simpleBonus = 0;
let deliveryCredit = 0;
let deliveryBonus = 0;
const arr = Object.keys(result).map(k => result[k])[0];
arr.splice(0, 5);
arr.splice(arr.length - 3, 3);

const simple = arr.filter(item => item.N !== 'Бонусы проведены в доставке');
const delivery = arr.filter(item => item.N === 'Бонусы проведены в доставке');

simple.forEach(item => simpleCredit = item.I + simpleCredit);
simple.forEach(item => simpleBonus = item.J + simpleBonus);
delivery.forEach(item => deliveryCredit = item.I + deliveryCredit);
delivery.forEach(item => deliveryBonus = item.J + deliveryBonus);

const allSums = {
  'Стандарт Кредит договор': simpleCredit,
  'Стандарт Бонус': simpleBonus,
  'Доставка Кредит договор': deliveryCredit,
  'Доставка Бонус': deliveryBonus
}

console.log(allSums);

