const express = require('express')
const exl = require('exceljs')
const fs = require('fs')
const cio = require('cheerio')
const app = express()
const { convertCsvToXlsx } = require ('@aternus/csv-to-xlsx')

app.use('/home', express.static(__dirname + '/public/index.html'))

app.get('/ozon', async function(req, res){

    const nat_cat = []
    const new_orders = []
    const difference = []
    const vendorCodes = []

    const colors = ['БЕЖЕВЫЙ', 'БЕЛЫЙ', 'БИРЮЗОВЫЙ', 'БОРДОВЫЙ', 'БРОНЗОВЫЙ', 'ВАНИЛЬ', 'ВИШНЯ', 'ГОЛУБОЙ', 'ЖЁЛТЫЙ', 'ЗЕЛЁНЫЙ', 'ЗОЛОТОЙ', 'ИЗУМРУДНЫЙ',
    'КАПУЧИНО', 'КИРПИЧНЫЙ', 'КОРАЛЛОВЫЙ', 'КОРИЧНЕВЫЙ', 'КРАСНЫЙ', 'ЛАЙМ', 'ЛЕОПАРД', 'МАЛИНОВЫЙ', 'МЕДНЫЙ', 'МОЛОЧНЫЙ', 'МЯТНЫЙ', 'ОЛИВКОВЫЙ', 'ОРАНЖЕВЫЙ',
    'ПЕСОЧНЫЙ', 'ПЕРСИКОВЫЙ', 'ПУРПУРНЫЙ', 'РАЗНОЦВЕТНЫЙ', 'РОЗОВО-БЕЖЕВЫЙ', 'РОЗОВЫЙ', 'СВЕТЛО-БЕЖЕВЫЙ', 'СВЕТЛО-ЗЕЛЕНЫЙ', 'СВЕТЛО-КОРИЧНЕВЫЙ', 'СВЕТЛО-РОЗОВЫЙ',
    'СВЕТЛО-СЕРЫЙ', 'СВЕТЛО-СИНИЙ', 'СВЕТЛО-ФИОЛЕТОВЫЙ', 'СЕРЕБРЯНЫЙ', 'СЕРО-ЖЕЛТЫЙ', 'СЕРО-ГОЛУБОЙ', 'СЕРЫЙ', 'СИНИЙ', 'СИРЕНЕВЫЙ', 'ЛИЛОВЫЙ', 'СЛИВОВЫЙ',
    'ТЕМНО-БЕЖЕВЫЙ', 'ТЕМНО-ЗЕЛЕНЫЙ', 'ТЕМНО-КОРИЧНЕВЫЙ', 'ТЕМНО-РОЗОВЫЙ', 'ТЕМНО-СЕРЫЙ', 'ТЕМНО-СИНИЙ', 'ТЕМНО-ФИОЛЕТОВЫЙ', 'ТЕРРАКОТОВЫЙ', 'ФИОЛЕТОВЫЙ',
    'ФУКСИЯ', 'ХАКИ', 'ЧЕРНЫЙ', 'ШОКОЛАДНЫЙ'
    ]
    
    const filePath = './public/new_orders/new_orders.html'

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    async function createImport(new_products) {
        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        const spans = content('span')

        spans.each((i, elem) => {
            if(content(elem).text().indexOf('00-') === 0) {
                if(new_products.includes((content(elem.parentNode.nextSibling).text()).trim())) {
                    vendorCodes.push(content(elem).text().replace(',', ''))
                    if(content(elem.parentNode.nextSibling).text().indexOf('Белый') >= 0) colors.push('белый')
                }
            }
        })

        for(i = 0; i < new_products.length; i++) {
            let size = ''            
                ws.getCell(`A${cellNumber}`).value = '6302'
                ws.getCell(`B${cellNumber}`).value = new_products[i]
                ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
                ws.getCell(`D${cellNumber}`).value = 'Артикул'
                ws.getCell(`E${cellNumber}`).value = vendorCodes[i]
                for(let c = 0; c < colors.length; c++) {
                    str = colors[c].toLowerCase()
                    elem = new_products[i].toLowerCase()
                    if(elem.indexOf(str) >= 0) {
                        ws.getCell(`G${cellNumber}`).value = colors[c].toUpperCase()
                    }
                }
                ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

                if(new_products[i].indexOf('Постельное') >= 0 || new_products[i].indexOf('Детское') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'КОМПЛЕКТ'
                }
                
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

                //Вставка размера начало
                
                if(new_products[i].indexOf('Постельное') >= 0) {
                    if(new_products[i].indexOf('1,5 спальное') >= 0 || new_products[i].indexOf('1,5 спальный') >= 0) {
                        size = '1,5 спальное'
                        if(new_products[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(new_products[i].indexOf('2 спальное') >= 0 || new_products[i].indexOf('2 спальный') >= 0) {
                        size = '2 спальное'
                        if(new_products[i].indexOf('с Евро') >= 0) {
                            size += ' с Евро простыней'
                        }
                        if(new_products[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(new_products[i].indexOf('Евро -') >= 0 || new_products[i].indexOf('евро -') >= 0 || new_products[i].indexOf('Евро на') >= 0 || new_products[i].indexOf('евро на') >= 0) {
                        size = 'Евро'
                        if(new_products[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(new_products[i].indexOf('Евро Макси') >= 0 || new_products[i].indexOf('евро макси') >= 0 || new_products[i].indexOf('Евро макси') >= 0) {
                        size = 'Евро Макси'
                        if(new_products[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(new_products[i].indexOf('семейное') >= 0 || new_products[i].indexOf('семейный') >= 0) {
                        size = 'семейное'
                        if(new_products[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(new_products[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                } else {
                    for(let k = 40; k < 305; k+=5) {
                        for(let l = 40; l < 305; l+=5) {
                            if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                size = `${k.toString()}х${l.toString()}`
                                for(let j = 10; j < 50; j+=10) {
                                    if(new_products[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || new_products[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                        size = `${k.toString()}х${l.toString()}х${j.toString()}`
                                        ws.getCell(`K${cellNumber}`).value = size
                                    } else {
                                        ws.getCell(`K${cellNumber}`).value = size
                                    }
                                }
                            }
                        }
                    }
                }
                
                //Вставка размера конец

                ws.getCell(`L${cellNumber}`).value = '6302100001'
                ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
                ws.getCell(`N${cellNumber}`).value = 'На модерации'                

                cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }

        // ws.mergeCells('D2:E2')

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}` : filePath = `./public/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    function getOrdersList(i, count) {
        if(count === 1) {
            const divs = content('.details-cell_propsSecond_f-KWL')            
            divs.each((i, elem) => {
                // console.log(content(elem).text())
                let str = (content(elem).text()).trim()
                if(str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
            })
        } else {
            for(i; i <= count; i++) {
                const divs = content('.details-cell_propsSecond_f-KWL')
                divs.each((i, elem) => {
                    // console.log(content(elem).text())
                    let str = (content(elem).text()).trim()
                    if(str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
                })  
            }
        }
    }

    getOrdersList(1, 1)

    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'

    let html = ''

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
            html += `<p>${elem}</p>`
        })

        // html = '<h1 class="success">Import successfully done</h1>'
        res.send(html)

        createImport(difference)
        // updateImport(nat_cat)

    }).catch(err => {
        console.log(err.message)
    })    

})

app.get('/wildberries', async function(req, res){
    
    const difference = []
    const wb_orders = []
    const nat_cat = []
    const vendors = []
    const names = []
    const ozon = []

    const colors = ['БЕЖЕВЫЙ', 'БЕЛЫЙ', 'БИРЮЗОВЫЙ', 'БОРДОВЫЙ', 'БРОНЗОВЫЙ', 'ВАНИЛЬ', 'ВИШНЯ', 'ГОЛУБОЙ', 'ЖЁЛТЫЙ', 'ЗЕЛЁНЫЙ', 'ЗОЛОТОЙ', 'ИЗУМРУДНЫЙ',
                        'КАПУЧИНО', 'КИРПИЧНЫЙ', 'КОРАЛЛОВЫЙ', 'КОРИЧНЕВЫЙ', 'КРАСНЫЙ', 'ЛАЙМ', 'ЛЕОПАРД', 'МАЛИНОВЫЙ', 'МЕДНЫЙ', 'МОЛОЧНЫЙ', 'МЯТНЫЙ', 'ОЛИВКОВЫЙ', 'ОРАНЖЕВЫЙ',
                        'ПЕСОЧНЫЙ', 'ПЕРСИКОВЫЙ', 'ПУРПУРНЫЙ', 'РАЗНОЦВЕТНЫЙ', 'РОЗОВО-БЕЖЕВЫЙ', 'РОЗОВЫЙ', 'СВЕТЛО-БЕЖЕВЫЙ', 'СВЕТЛО-ЗЕЛЕНЫЙ', 'СВЕТЛО-КОРИЧНЕВЫЙ', 'СВЕТЛО-РОЗОВЫЙ',
                        'СВЕТЛО-СЕРЫЙ', 'СВЕТЛО-СИНИЙ', 'СВЕТЛО-ФИОЛЕТОВЫЙ', 'СЕРЕБРЯНЫЙ', 'СЕРО-ЖЕЛТЫЙ', 'СЕРО-ГОЛУБОЙ', 'СЕРЫЙ', 'СИНИЙ', 'СИРЕНЕВЫЙ', 'ЛИЛОВЫЙ', 'СЛИВОВЫЙ',
                        'ТЕМНО-БЕЖЕВЫЙ', 'ТЕМНО-ЗЕЛЕНЫЙ', 'ТЕМНО-КОРИЧНЕВЫЙ', 'ТЕМНО-РОЗОВЫЙ', 'ТЕМНО-СЕРЫЙ', 'ТЕМНО-СИНИЙ', 'ТЕМНО-ФИОЛЕТОВЫЙ', 'ТЕРРАКОТОВЫЙ', 'ФИОЛЕТОВЫЙ',
                        'ФУКСИЯ', 'ХАКИ', 'ЧЕРНЫЙ', 'ШОКОЛАДНЫЙ'
                        ]

    const wb = new exl.Workbook()

    const hsFile = './public/Краткий отчет.xlsx'
    const ozonFile = './public/products.xlsx'
    const wbFile = './public/wildberries/new.xlsx'

    let html = ''

    await wb.xlsx.readFile(hsFile)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    await wb.xlsx.readFile(wbFile)

    const _ws = wb.getWorksheet('Сборочные задания')

    const c12 = _ws.getColumn(12)

    c12.eachCell(c => {
        wb_orders.push(c.value)
    })

    await wb.xlsx.readFile(ozonFile)

    const ws_ = wb.getWorksheet('Worksheet')

    const c1 = ws_.getColumn(1)
    const c6 = ws_.getColumn(6)

    c1.eachCell(c => {
        vendors.push(c.value.replace(`'`,``))
    })

    c6.eachCell(c => {
        names.push(c.value.trim())
    })

    wb_orders.forEach(elem => {
        if(vendors.indexOf(elem) >= 0){
            let index = vendors.indexOf(elem)
            ozon.push(names[index])
            // console.log(typeof index)
        }
    })

    ozon.forEach(elem => {
        if(nat_cat.indexOf(elem) < 0) {
            difference.push(elem)
        }
    })

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(i = 0; i < array.length; i++) {
            let size = ''            
                ws.getCell(`A${cellNumber}`).value = '6302'
                ws.getCell(`B${cellNumber}`).value = array[i]
                ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
                ws.getCell(`D${cellNumber}`).value = 'Артикул'
                let index = names.indexOf(array[i])
                ws.getCell(`E${cellNumber}`).value = vendors[index]
                for(let c = 0; c < colors.length; c++) {
                    str = colors[c].toLowerCase()
                    elem = array[i].toLowerCase()
                    if(elem.indexOf(str) >= 0) {
                        ws.getCell(`G${cellNumber}`).value = colors[c].toUpperCase()
                    }
                }
                ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

                if(array[i].indexOf('Постельное') >= 0 || array[i].indexOf('Детское') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'КОМПЛЕКТ'
                }
                
                if(array[i].indexOf('Простыня') >= 0) {
                    if(array[i].indexOf('на резинке') >= 0) {
                        ws.getCell(`F${cellNumber}`).value = 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                    } else {
                        ws.getCell(`F${cellNumber}`).value = 'ПРОСТЫНЯ'
                    }
                }
                if(array[i].indexOf('Пододеяльник') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
                }
                if(array[i].indexOf('Наволочка') >= 0) {
                    if(array[i].indexOf('50х70') >=0 || array[i].indexOf('40х60') >= 0 || array[i].indexOf('50 х 70') >=0 || array[i].indexOf('40 х 60') >= 0) {
                        ws.getCell(`F${cellNumber}`).value = 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                    } else {
                        ws.getCell(`F${cellNumber}`).value = 'НАВОЛОЧКА КВАДРАТНАЯ'
                    }
                }
                if(array[i].indexOf('Наматрасник') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'НАМАТРАСНИК'
                }
                if(array[i].indexOf('страйп-сатин') >= 0 || array[i].indexOf('страйп сатин') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
                }
                if(array[i].indexOf('твил-сатин') >= 0 || array[i].indexOf('твил сатин') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ТВИЛ-САТИН'
                }
                if(array[i].indexOf('тенсел') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ТЕНСЕЛЬ'
                }
                if(array[i].indexOf('бяз') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'БЯЗЬ'
                }
                if(array[i].indexOf('поплин') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ПОПЛИН'
                }
                if(array[i].indexOf('сатин') >= 0 && array[i].indexOf('-сатин') < 0 && array[i].indexOf('п сатин') < 0 && array[i].indexOf('л сатин') < 0 && array[i].indexOf('сатин-') < 0 && array[i].indexOf('сатин ж') < 0) {
                    ws.getCell(`I${cellNumber}`).value = 'САТИН'
                }
                if(array[i].indexOf('вареный') >= 0 || array[i].indexOf('варёный') >= 0 || array[i].indexOf('вареного') >= 0 || array[i].indexOf('варёного') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
                }
                if(array[i].indexOf('сатин-жаккард') >= 0 || array[i].indexOf('сатин жаккард') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'САТИН-ЖАККАРД'
                }
                if(array[i].indexOf('страйп-микрофибр') >= 0 || array[i].indexOf('страйп микрофибр') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'МИКРОФИБРА'
                }
                if(array[i].indexOf('шерст') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ПОЛИЭФИР'
                }

                if(array[i].indexOf('тенсел') >= 0) {ws.getCell(`J${cellNumber}`).value = '100% Эвкалипт'}
                else if(array[i].indexOf('шерст') >= 0) {ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'}
                else {ws.getCell(`J${cellNumber}`).value = '100% Хлопок'}

                //Вставка размера начало
                
                if(array[i].indexOf('Постельное') >= 0) {
                    if(array[i].indexOf('1,5 спальное') >= 0 || array[i].indexOf('1,5 спальный') >= 0) {
                        size = '1,5 спальное'
                        if(array[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(array[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(array[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(array[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(array[i].indexOf('2 спальное') >= 0 || array[i].indexOf('2 спальный') >= 0) {
                        size = '2 спальное'
                        if(array[i].indexOf('с Евро') >= 0) {
                            size += ' с Евро простыней'
                        }
                        if(array[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(array[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(array[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(array[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(array[i].indexOf('Евро -') >= 0 || new_products[i].indexOf('евро -') >= 0 || array[i].indexOf('Евро на') >= 0 || array[i].indexOf('евро на') >= 0) {
                        size = 'Евро'
                        if(array[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(array[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(array[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(array[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(array[i].indexOf('Евро Макси') >= 0 || array[i].indexOf('евро макси') >= 0 || array[i].indexOf('Евро макси') >= 0) {
                        size = 'Евро Макси'
                        if(array[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(array[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(array[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(array[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                    if(array[i].indexOf('семейное') >= 0 || array[i].indexOf('семейный') >= 0) {
                        size = 'семейное'
                        if(array[i].indexOf('на резинке') >= 0) {
                            size += ' на резинке'
                            for(let k = 40; k < 305; k+=5) {
                                for(let l = 40; l < 305; l+=5) {
                                    if(array[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                        for(let j = 10; j < 50; j+=10) {
                                            if(array[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                                size += ` ${k.toString()}х${l.toString()}х${j.toString()}`
                                                ws.getCell(`K${cellNumber}`).value = size
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if(array[i].indexOf('с наволочками 50х70') >= 0) {
                            size += ' с наволочками 50х70'
                            ws.getCell(`K${cellNumber}`).value = size
                        } else {
                            ws.getCell(`K${cellNumber}`).value = size
                        }
                    }
                } else {
                    for(let k = 40; k < 305; k+=5) {
                        for(let l = 40; l < 305; l+=5) {
                            if(array[i].indexOf(` ${k.toString()}х${l.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()}`) >= 0) {
                                size = `${k.toString()}х${l.toString()}`
                                for(let j = 10; j < 50; j+=10) {
                                    if(array[i].indexOf(` ${k.toString()}х${l.toString()}х${j.toString()}`) >= 0 || array[i].indexOf(` ${k.toString()} х ${l.toString()} х ${j.toString()}`) >= 0) {
                                        size = `${k.toString()}х${l.toString()}х${j.toString()}`
                                        ws.getCell(`K${cellNumber}`).value = size
                                    } else {
                                        ws.getCell(`K${cellNumber}`).value = size
                                    }
                                }
                            }
                        }
                    }
                }
                
                //Вставка размера конец

                ws.getCell(`L${cellNumber}`).value = '6302100001'
                ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
                ws.getCell(`N${cellNumber}`).value = 'На модерации'                

                cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/wildberries/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}` : filePath = `./public/wildberries/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    createImport(difference)

    difference.forEach(elem => {
        html += `<p>${elem}</p>`
    })

    res.send(html)

})

app.listen(3030)