const listShipmentID = ["FBA16XJXW9Z6", "FBA16V77Q3T4", "FBA16PFFKY38", "FBA16P4CTPN2"];
const targetValue = "FBA16V77Q3T4";

const elementsAfter = listShipmentID.filter(value => value !== targetValue).slice(0, listShipmentID.indexOf(targetValue));

console.log(elementsAfter);
