import fs, { writeFileSync } from 'fs';
import { default as xlsx } from 'node-xlsx';
import { default as path } from 'path';

class Xlsx {
    constructor() {
        this.file = null;
        this.data = null;
        this.worksheet = null;
    }

    setOriginalFile(file) {
        this.file = `./${file}`;
    }

    setData(data) {
        this.data = data;
        if (this.file && typeof this.data === 'object' && Object.keys(this.data).length > 0) {
            this.worksheet = xlsx.parse(fs.readFileSync(this.file))[0];
            const sheetData = this.worksheet.data;
    
            this.data.forEach(({ c, r, v }) => {
                const columnIndex = c.charCodeAt(0) - 65;
                if (sheetData[r - 1] && sheetData[r - 1][columnIndex] !== undefined) {
                    sheetData[r - 1][columnIndex] = v;
                } else {
                    console.log(`A célula ${c}${r} está fora dos limites.`);
                }
            });
        }
    }
    

    download() {
        if (!this.file || !this.worksheet) {
            console.log('File or worksheet not specified.');
            return;
        }

        const { name, ext,dir } = path.parse(this.file);
        const fileName = `${dir}/${name}_transf${ext}`;
        writeFileSync(fileName, xlsx.build([this.worksheet]));

        return fileName;
    }
}

export default Xlsx;
