const fs = require('fs')
const csv = require('csv-parser')
const xlsx = require('node-xlsx')

// Excel sucks when dealing with CSV value contains line breaks

const results = []
async function makeExcelHappy() {
    const fileName = process.argv[2]
    if (fileName && fileName.indexOf('.csv') > -1 ) {
        fs.createReadStream(fileName)
            .pipe(csv())
            .on('data', (data) => {
                results.push(Object.keys(data).map(key => data[key]))
            })
            .on('end', () => {
                console.log('start writing xls file')
                const buffer = xlsx.build([{name: 'csv2xls', data: results}])
                const fileNameWithoutExt = fileName.replace('.csv', '')

                fs.writeFile(fileNameWithoutExt + '.xlsx', buffer, function (err) {
                    if (err) {
                        console.log('error while exporting to xlsx', err)
                    } else {
                        console.log('file exported successfully')
                    }
                })
            })
    } else {
        console.log('expect a csv file name as input, got `'+fileName + '`')
    }
}

makeExcelHappy()