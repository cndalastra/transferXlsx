import Xlsx from './class/Xlsx.js'

const fileName = 'files/Original.xlsx'
const newData = [
    { c: 'A', r: 1, v: 'New Value A1' },
    { c: 'B', r: 2, v: 'New Value B2' },
    { c: 'G', r: 2, v: 'New Value' },
    { c: 'E', r: 2, v: 'New Value' },
    { c: 'R', r: 2, v: 'New Value' }
]



const xlsx = new Xlsx();

xlsx.setOriginalFile(fileName);
xlsx.setData(newData);
xlsx.download();