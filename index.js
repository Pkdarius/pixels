const fs = require('fs');
const getPixels = require('get-pixels');
const exceljs = require('exceljs');
const readline = require('readline');
const { Console } = require('console');

const numberToHexString = (number) => {
    const hex = number.toString(16);
    return hex.length === 1 ? `0${hex}`: hex;
}

const main = async () => {
    const arguments = process.argv.slice(2);
    const filePathInput = arguments[0];
    checkFileInput(filePathInput);

    const {width, height, imageColors} = await getImageDataArray(filePathInput);

    const fileName = getFileName(filePathInput);
    drawWorkbook(width, height, imageColors, fileName);
}

const checkFileInput = (path) => {
    console.log('Checking for file input...');
    if (!path.match(/.*\.(jpg|png)$/i)) {
        console.error('Only accept jpg or png file format');
        process.exit(1);
    }

    if(!fs.existsSync(path)) {
        console.error('File input does not exist');
        process.exit(2);
    }
}

const getFileName = (path) => {
    const fileInputPathSplitted = path.split('\\');
    const fileInputName = fileInputPathSplitted[fileInputPathSplitted.length - 1];
    return fileInputName.substring(0, fileInputName.lastIndexOf('.'));
}

const getImageDataArray = (path) => {
    return new Promise((resolve, reject) => {
        getPixels(path, async (err, pixels) => {
            console.log('Getting image data...');
            if (err) return console.error(err);
            const { data, shape } = pixels;
            const [width, height] = shape;
            const colors = [];
        
            for (let i = 0; i < data.length; i += 4) {
                const a = numberToHexString(data[i]);
                const b = numberToHexString(data[i + 1]);
                const c = numberToHexString(data[i + 2]);
                colors.push('ff' + a + b + c);
            }
            const imageColors = [];
            while(colors.length) {
                imageColors.push(colors.splice(0, width));
            }
            resolve({width, height, imageColors});            
        });
    });
}

const drawWorkbook = async (width, height, imageColors, fileName) => {
    console.log('Drawing...');
    const workbook = new exceljs.Workbook();
    const sheet = workbook.addWorksheet('Hi', {
        views: [
            {
                x: 0,
                y: 0, 
                width, 
                height, 
                zoomScale: 50,
                showGridLines: false
            }
        ]
    });

    for (let i = 0; i < height; i++) {
        let row = [];
        for(let j = 0; j < width; j++) {
            row.push(null)
        }
        const currentRow = sheet.addRow(row);
        currentRow.height = 4;
    }

    for (let i = 0; i < width; i++) {
        const column = sheet.getColumn(i + 1);
        column.width = 1;
    }

    for (let i = 0; i < height; i++) {
        const row = sheet.getRow(i);
        for(let j = 0; j < width; j++) {
            const cell = row.getCell(j + 1);
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor:{argb: imageColors[i][j]}
            }
        }
    }
    await workbook.xlsx.writeFile(`${fileName}.xlsx`);
    console.log('Done!');
}

main();