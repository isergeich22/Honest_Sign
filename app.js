const express = require('express')
const exl = require('exceljs')
const fs = require('fs')
const cio = require('cheerio')
const app = express()

app.use('/home', express.static(__dirname + '/public'))

app.get('/excel', async function(req, res){

    const nat_cat = []
    const new_orders = []
    const difference = []
    const rows = []
    const vendorCodes = []

    // const _wb = new exl.Workbook()

    // console.log(_wb)

    // await _wb.xlsx.readFile('./public/IMPORT_TNVED_6302 (3).xlsx')

    // const _ws = _wb.getWorksheet('IMPORT_TNVED_6302')

    // const _r2 = _ws.getRow(2)

    // _r2.eachCell((cell, cn) => {
    //     console.log(cell.value)
    // })

    // const _r5 = _ws.getRow(5)

    // _r5.eachCell((cell, cn) => {
    //     cell.value = 'Test'
    // })

    // await _wb.xlsx.writeFile('IMPORT_TNVED_6302_16_03.xlsx')

    function createImport(new_products) {
        const fileName = './public/IMPORT.xlsx'

        const wb = new exl.Workbook()
        const ws = wb.addWorksheet('IMPORT')        

        let cellNumber = 5

        const r1 = ws.getRow(1)
        r1.values = ['Код ТНВЭД', 'Полное наименование товара', 'Товарный знак', 'Модель / артикул производителя', '', 'Вид товара', 'Цвет', 'Возраст потребителя', 'Тип текстиля', 'Состав', 'Размер изделия', 'Код ТНВЭД', 'Номер Регламента/стандарта', 'Статус карточки товара в Каталоге', 'Результат обработки данных в Каталоге']
        r1.eachCell((cell, cn) => {
            cell.alignment = {vertical: 'bottom', horizontal: 'center'}
        })
        r1.font = {name: 'Arial', size: 10, bold: true}

        r1.eachCell((cell, cn) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor:{argb: 'FFFB00'}
            }
        })

        r1.getCell(15).fill = {
            type:'pattern',
            pattern:'solid',
            fgColor:{argb:'E3E3E3'}
        }

        r1.getCell(15).font = {
            name: 'Arial',
            bold: false,
            size: 10
        }

        const rows = ws.getRows(2, 348)

        rows.forEach(el => {
            el.font = {name: 'Arial', size: 10}
        })

        ws.mergeCells('D1:E1')

        const r2 = ws.getRow(2)
        r2.values = ['Tnved', '2478', '2504', '13914', '13914', '12', '36', '557', '13967', '2483', '15435', '13933', '13836', 'status','result']
        ws.mergeCells('D2:E2')
        r2.eachCell((cell, cn) => {
            cell.alignment = {vertical: 'bottom', horizontal: 'center'}
        })
        r2.eachCell((cell, cn) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor:{argb:'E3E3E3'}
            }
        })

        const r3 = ws.getRow(3)
        r3.values = ['value', 'value', 'value', 'type', 'value', 'value', 'value', 'value', 'value', 'value', 'value', 'value', 'value', 'value', 'value']
        r3.eachCell((cell, cn) => {
            cell.alignment = {vertical: 'bottom', horizontal: 'center'}
        })
        r3.eachCell((cell, cn) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor:{argb:'E3E3E3'}
            }
        })

        const r4 = ws.getRow(4)
        r4.values = ['', 'Текстовое значение', 'Значение из справочника, Текстовое значение', 'Тип (из справочника)', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое значение', 'Текстовое поле (Черновик или На модерации)', 'Заполняется автоматически при загрузке в систему']
        r4.eachCell((cell, cn) => {
            cell.alignment = {vertical: 'bottom', horizontal: 'center'}
        })
        r4.eachCell((cell, cn) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor:{argb:'E3E3E3'}
            }
        })

        r4.getCell(15).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'FFC096'}
        }
               
        for(i = 0; i < new_products.length; i++) {
            if(new_products[i].indexOf('Постельное') < 0) {
                ws.getCell(`A${cellNumber}`).value = '6302'
                ws.getCell(`B${cellNumber}`).value = new_products[i]
                ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
                ws.getCell(`D${cellNumber}`).value = 'Артикул'
                ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'
                if(new_products[i].indexOf('Простыня') >= 0) {
                    if(new_products[i].indexOf('на резинке') >= 0) {
                        ws.getCell(`F${cellNumber}`).value = 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                    } else {
                        ws.getCell(`F${cellNumber}`).value = 'ПРОСТЫНЯ'
                    }
                }
                if(new_products[i].indexOf('Пододеяльник') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
                }
                if(new_products[i].indexOf('Наволочка') >= 0) {
                    if(new_products[i].indexOf('50х70') >=0 || new_products[i].indexOf('40х60') >= 0 || new_products[i].indexOf('50 х 70') >=0 || new_products[i].indexOf('40 х 60') >= 0) {
                        ws.getCell(`F${cellNumber}`).value = 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                    } else {
                        ws.getCell(`F${cellNumber}`).value = 'НАВОЛОЧКА КВАДРАТНАЯ'
                    }
                }
                if(new_products[i].indexOf('Наматрасник') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'НАМАТРАСНИК'
                }
                if(new_products[i].indexOf('страйп-сатин') >= 0 || new_products[i].indexOf('страйп сатин') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
                }
                if(new_products[i].indexOf('твил-сатин') >= 0 || new_products[i].indexOf('твил сатин') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ТВИЛ-САТИН'
                }
                if(new_products[i].indexOf('тенсел') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ТЕНСЕЛЬ'
                }
                if(new_products[i].indexOf('бяз') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'БЯЗЬ'
                }
                if(new_products[i].indexOf('поплин') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ПОПЛИН'
                }
                if(new_products[i].indexOf('сатин') >= 0 && new_products[i].indexOf('-сатин') < 0 && new_products[i].indexOf('п сатин') < 0 && new_products[i].indexOf('л сатин') < 0 && new_products[i].indexOf('сатин-') < 0 && new_products[i].indexOf('сатин ж') < 0) {
                    ws.getCell(`I${cellNumber}`).value = 'САТИН'
                }
                if(new_products[i].indexOf('вареный') >= 0 || new_products[i].indexOf('варёный') >= 0 || new_products[i].indexOf('вареного') >= 0 || new_products[i].indexOf('варёного') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
                }
                if(new_products[i].indexOf('сатин-жаккард') >= 0 || new_products[i].indexOf('сатин жаккард') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'САТИН-ЖАККАРД'
                }
                if(new_products[i].indexOf('страйп-микрофибр') >= 0 || new_products[i].indexOf('страйп микрофибр') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'МИКРОФИБРА'
                }
                if(new_products[i].indexOf('шерст') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ПОЛИЭФИР'
                }

                if(new_products[i].indexOf('тенсел') >= 0) {ws.getCell(`J${cellNumber}`).value = '100% Эвкалипт'}
                else if(new_products[i].indexOf('шерст') >= 0) {ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'}
                else {ws.getCell(`J${cellNumber}`).value = '100% Хлопок'}

                ws.getCell(`L${cellNumber}`).value = '6302100001'
                ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
                ws.getCell(`N${cellNumber}`).value = 'На модерации'                

                cellNumber++
            }
        }

        wb.xlsx
            .writeFile(fileName)
            .then(() => {
                console.log('File refresh: success')
            })
            .catch(err => {
                console.log(err.message)
            })

    }

    function updateImport(new_products) {
        const fileName = './public/IMPORT.xlsx'

        let cellNumber = 5

        const wb = new exl.Workbook()
        wb.xlsx.readFile(fileName).then(() => {
            wb.eachSheet((ws, sheetId) => {
                for(i = 0; i < new_products.length; i++) {
                    ws.getCell(`B${cellNumber}`).value = new_products[i]
                    cellNumber++
                }
            })
        })

        let date_ob = new Date()

        console.log(`${(date_ob.getDate()).toString()}_${(date_ob.getMonth()).toString}`)

        wb.xlsx
            .writeFile(`./public/IMPORT-new.xlsx`)
            .then(() => {
                console.log('File updated successfully')
            })
            .catch(err => {
                console.log(err.message)
            })

    }

    function getOrdersList(i, count) {
        if(count === 1) {
            const filePath = './public/new_orders/new_orders.html'

            const fileContent = fs.readFileSync(filePath, 'utf-8')

            const content = cio.load(fileContent)
            const spans = content('span')
            const divs = content('.details-cell_propsSecond_f-KWL')
            // console.log(spans)
            spans.each((i, elem) => {
                let str = (content(elem).text()).replace(',', '')
                if(str.indexOf('00-') >= 0) vendorCodes.push(str)
            })
            // console.log(vendorCodes)
            divs.each((i, elem) => {
                // console.log(content(elem).text())
                let str = (content(elem).text()).trim()
                if(str.indexOf('Постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
            })
        } else {
            for(i; i <= count; i++) {
                const filePath = `./public/new_orders/new_orders_${i}.html`

                const fileContent = fs.readFileSync(filePath, 'utf-8')
    
                const content = cio.load(fileContent)
                const divs = content('.details-cell_propsSecond_f-KWL')
                divs.each((i, elem) => {
                    // console.log(content(elem).text())
                    let str = (content(elem).text()).trim()
                    if(str.indexOf('Постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
                })  
            }
        }
    }

    // let files = fs.readdir('/new_orders')

    getOrdersList(1, 1)

    let html = ''
    // let _html = ''
    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'

    wb.xlsx.readFile(fileName).then(() => {
        
        const ws = wb.getWorksheet('Краткий отчет')

        const c2 = ws.getColumn(2)

        c2.eachCell(c => {
           nat_cat.push(c.value)
        })

        for(i = 0; i < new_orders.length; i++) {
            if(nat_cat.indexOf(new_orders[i]) < 0 && difference.indexOf(new_orders[i]) < 0){
                difference.push(new_orders[i])
            }
        }

        difference.forEach(elem => {
            if(elem.indexOf('Постельное') < 0) html += `<p>${elem}</p>`
        })

        // html = '<h1 class="success">Import successfully done</h1>'
        res.send(html)

        createImport(difference)
        // updateImport(nat_cat)

    }).catch(err => {
        console.log(err.message)
    })    

})

app.listen(3030)