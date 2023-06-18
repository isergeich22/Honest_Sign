const express = require('express')
const exl = require('exceljs')
const fs = require('fs')
const cio = require('cheerio')
const app = express()

const headerComponent = `<!DOCTYPE html>
                            <html lang="en">
                            <head>
                                <meta charset="UTF-8">
                                <meta http-equiv="X-UA-Compatible" content="IE=edge">
                                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                <link rel="stylesheet" href="/css/styles.css" type="text/css">
                                <link rel="shortcut icon" type="image/png" href="/favicon.png">`

const navComponent = `<header class="header">
                        <nav>
                            <img src="/img/chestnyj_znak.png" alt="честный знак">
                            <p class="nav-item" id="home">Главная</p>
                            <p class="nav-item" id="import">Создание импорт-файлов</p>
                            <p class="nav-item" id="cis_actions">Действия с КИЗ</p>                        
                        </nav>                    
                    </header>`

const footerComponent = `   <button id="top" class="button-top">
                            <svg width="24" height="24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                <g clip-path="url(#ArrowLongUp_large_svg__clip0_35331_5070)">
                                    <path d="M12 2v20m0-20l7 6.364M12 2L5 8.364" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>
                                </g><defs><clipPath id="ArrowLongUp_large_svg__clip0_35331_5070"><path fill="#fff" transform="rotate(90 12 12)" d="M0 0h24v24H0z">
                                </path></clipPath></defs></svg>
                            </button>    
                            <script src="/script.js"></script>
                            </body>
                        </html>`

app.use(express.static(__dirname + '/public'))

app.get('/home', async function(req, res){

    let html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

    async function getNationalCatalog() {
        
        const names = []
        const gtins = []

        const wb = new exl.Workbook()

        await wb.xlsx.readFile('./public/Краткий отчет.xlsx')

        const ws = wb.getWorksheet('Краткий отчет')

        const [c1, c2] = [ws.getColumn(1), ws.getColumn(2)]

        c1.eachCell(c => {
            gtins.push(`0${c.value}`)
        })

        c2.eachCell(c => {
            names.push(c.value)
        })

        return [names, gtins]

    }

    async function getMonthlyMarks() {

        const actual_gtins = []
        const actual_marks = []
        const actual_dates = []

        const wb = new exl.Workbook()

        await wb.xlsx.readFile('./public/actual_marks.xlsx')

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c2, c23] = [ws.getColumn(1), ws.getColumn(2), ws.getColumn(23)]

        c1.eachCell(c => {
            if(c.value.indexOf('01') >= 0) {
                let str = c.value
                if(str.indexOf('<') >= 0) {
                    str = str.replace(/</g, '&lt;')                    
                }
                actual_marks.push(str)
                
            }
        })

        c2.eachCell(c => {
            if(c.value !== null) {
                if(c.value.indexOf('029') >= 0) {
                    actual_gtins.push(c.value)
                }
            }
        })

        c23.eachCell(c => {
            if(c.value !== null) {
                if(c.value.indexOf('-') >= 0) {
                    let str = c.value
                    actual_dates.push(str.replace(str.substring(10), ''))
                }
            }
        })

        return [actual_gtins, actual_marks, actual_dates]

    }

    async function renderMarksTable() {
        
        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates] = await getMonthlyMarks()

        html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">КИЗ</div>
                            <div class="header-cell">GTIN</div>
                            <div class="header-cell">Товар</div>
                            <div class="header-cell">Дата эмиссии</div>
                            <div class="header-cell">Статус</div>
                        </div>
                        <div class="header-wrapper"></div>`

        for(let i = 0; i < actual_marks.length; i++) {
            html+= `<div class="table-row">
                        <span type="text" id="mark">${actual_marks[i]}</span>
                        <span id="gtin">${actual_gtins[i]}</span>
                        <span id="name">${names[gtins.indexOf(actual_gtins[i])]}</span>
                        <span id="date">${actual_dates[i]}</span>
                        <span id="status">В обороте</span>
                    </div>`
        }

        

    }

    await renderMarksTable()

    html += `</div>
        </section>
    ${footerComponent}`

    res.send(html)
})

app.get('/ozon', async function(req, res){

    let html = `${headerComponent}
                    <title>Импорт - OZON</title>
                </head>
                <body>
                        ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

    const nat_cat = []
    const new_orders = []
    const new_items = []
    const current_items = []
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

                if(new_products[i].indexOf('Полотенце') >= 0) {
                    ws.getCell(`F${cellNumber}`).value = 'ИЗДЕЛИЯ ДЛЯ САУНЫ'
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
                if(new_products[i].indexOf('перкал') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'ПЕРКАЛЬ'
                }
                if(new_products[i].indexOf('махра') >= 0 || new_products[i].indexOf('махровое') >= 0) {
                    ws.getCell(`I${cellNumber}`).value = 'МАХРОВАЯ ТКАНЬ'
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

        month < 10 ? filePath = `./public/ozon/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_ozon` : filePath = `./public/ozon/IMPORT_TNVED_6302_${date_ob.getDate()}_${month}_ozon`

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
                if(str.indexOf('Полотенце') >= 0 || str.indexOf('полотенце') >= 0 || str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
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

    wb.xlsx.readFile(fileName).then(() => {
        
        const ws = wb.getWorksheet('Краткий отчет')

        const c2 = ws.getColumn(2)

        c2.eachCell(c => {
           nat_cat.push(c.value)
        })

        for(i = 0; i < new_orders.length; i++) {
            if(nat_cat.indexOf(new_orders[i]) < 0 && new_items.indexOf(new_orders[i]) < 0){

                new_items.push(new_orders[i])

            }

            if(nat_cat.indexOf(new_orders[i]) >=0 && current_items.indexOf(new_orders[i]) < 0) {

                current_items.push(new_orders[i])

            }
        }

        html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Наименование</div>
                            <div class="header-cell">Статус</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

        new_items.forEach(elem => {
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-new">Новый товар</span>
                     </div>`
        })

        current_items.forEach(elem => {
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-current">Актуальный товар</span>
                     </div>`
        })

        html += `       </section>
                        <section class="action-form">
                            <button id="current-order"><a href="http://localhost:3030/ozon_marks_order">Создать заказ маркировки для актуальных товаров</a></button>
                            <button id="new-order"><a href="http://localhost:3030/ozon_new_marks_order" target="_blank">Создать заказ маркировки для новых товаров</a></button>
                        </section>
                        <div class="body-wrapper"></div>                        
                        ${footerComponent}`

        // html = '<h1 class="success">Import successfully done</h1>'
        res.send(html)

        createImport(new_items)

        }).catch(err => {
        console.log(err.message)
    })    

})

app.get('/ozon_marks_order', async function(req, res){
    
    const nat_cat = []
    const gtins = []
    const new_orders = []
    const current_items = []
    const current_quantity = []
    const quantity = []

    let html = `${headerComponent}
                    <title>Заказ акутальных маркировок</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

    const filePath = './public/new_orders/new_orders.html'

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    function getOrdersList(i, count) {
        if(count === 1) {
            const divs = content('.details-cell_propsSecond_f-KWL')            
            divs.each((i, elem) => {
                // console.log(content(elem).text())
                let str = (content(elem).text()).trim()                
                if(str.indexOf('Полотенце') >= 0 || str.indexOf('полотенце') >= 0 || str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
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

    function getQuantity() {
        const spans = content('.mr2')

        spans.each((i, elem) => {
            if(content(elem).text().indexOf('шт.') >= 0) {
                // if(new_orders.indexOf(content(elem.parentNode.nextSibling).text().trim()) >= 0) {
                //     quantity.push(parseInt((content(elem).text().replace(' шт.', ''))))
                // }
                if(new_orders.indexOf(content(elem.parentNode.nextSibling).text().trim()) >= 0) {
                    quantity.push(parseInt((content(elem).text().replace(' шт.', ''))))
                }
            }
        })
        // spans.each((i, elem) => {
        //     if(content(elem).text().indexOf('шт.') === 0) quantity.push(content(elem).text())
        // })
    }

    getOrdersList(1,1)

    // console.log(new_orders.length)

    getQuantity()

    // console.log(quantity.length)

    

    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'    

    await wb.xlsx.readFile(fileName)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c_1 = ws.getColumn(1)

    c_1.eachCell(c => {
        gtins.push(c.value)        
    })

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    // console.log(nat_cat)

    new_orders.forEach(elem => {
        if(nat_cat.indexOf(elem) >= 0 && current_items.indexOf(elem) < 0) {
            let index = new_orders.indexOf(elem)
            current_items.push(elem)
            current_quantity.push(quantity[index])
        } else {
            let index = new_orders.indexOf(elem)
            let i = current_items.indexOf(elem)
            current_quantity[i] += parseInt(quantity[index])
        }
    })

    html += `<div class="new_items_order">
                <h3>Товары не в заказе</h3><hr>`

    new_orders.forEach(elem => {
        if(nat_cat.indexOf(elem) < 0) {
            let index = new_orders.indexOf(elem)
            html += `<p class="new">${elem} - <span>${quantity[index]} шт.</span></p>`
        }
    })

    html += `<hr><button><a href="http://localhost:3030/ozon_new_marks_order" target="_blank">Создать заказы для товаров на модерации</a></button></div><div class="current_items_order">
                <h3>Товары в заказе</h3><hr>`

    for(let i = 0; i < current_items.length; i++) {
        html += `<p class="current">${current_items[i]} - <span>${current_quantity[i]} шт.</span></p>`
    }

    html += `   <hr></div>
            </section>
        ${footerComponent}`

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < current_items.length; i++) {

            _temp.push(current_items[i])
            
                if(_temp.length%10 === 0) {
                    orderList.push(_temp)
                    _temp = []
                }
        }        

        orderList.push(_temp)
        _temp = []

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < current_quantity.length; i++) {

            temp.push(current_quantity[i])

                if(temp.length%10 === 0) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                    temp = []
                }

        }

        quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        return quantityList

    }

    // console.log(createNameList())
    // console.log(createQuantityList())

    function createOrder() {        

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        for(let i = 0; i < List.length; i++) {
             content += `<?xml version="1.0" encoding="utf-8"?>
                        <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                            <lp>
                                <productGroup>lp</productGroup>
                                <contactPerson>333</contactPerson>
                                <releaseMethodType>REMARK</releaseMethodType>
                                <createMethodType>SELF_MADE</createMethodType>
                                <productionOrderId>OZON</productionOrderId>
                                <products>`
            for(let j = 0; j < List[i].length; j++) {                
                if(nat_cat.indexOf(List[i][j]) >= 0) {
                    content += `<product>
                                    <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                    <quantity>${Quantity[i][j]}</quantity>
                                    <serialNumberType>OPERATOR</serialNumberType>
                                    <cisType>UNIT</cisType>
                                    <templateId>10</templateId>
                                </product>`
                }
            }

            content += `    </products>
                        </lp>
                    </order>`

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_${month}.xml`

            fs.writeFileSync(filePath, content)

            content = ``

        }   

    }

    createOrder()

    res.send(html)

})

app.get('/ozon_new_marks_order', async function(req, res){

    const wb = new exl.Workbook()

    const new_orders = []
    const quantity = []
    const nat_cat = []
    const products = []
    const moderation_products = []
    const moderation_gtins = []
    const new_items = []
    const new_quantity = []

    let html = `${headerComponent}
                    <title>Заказ новых маркировок</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

    const hsFile = './public/Краткий отчет.xlsx'
    const filePath = './public/moderation_marks/moderation_marks.html'
    const filePathOzon = './public/new_orders/new_orders.html'

    const fileContentOzon = fs.readFileSync(filePathOzon, 'utf-8')

    const contentOzon = cio.load(fileContentOzon)

    function getOrdersList(i, count) {
        if(count === 1) {
            const divs = contentOzon('.details-cell_propsSecond_f-KWL')            
            divs.each((i, elem) => {
                // console.log(content(elem).text())
                let str = (contentOzon(elem).text()).trim()                
                if(str.indexOf('Полотенце') >= 0 || str.indexOf('полотенце') >= 0 || str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
            })
        } else {
            for(i; i <= count; i++) {
                const divs = contentOzon('.details-cell_propsSecond_f-KWL')
                divs.each((i, elem) => {
                    // console.log(content(elem).text())
                    let str = (contentOzon(elem).text()).trim()
                    if(str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
                })  
            }
        }
    }

    function getQuantity() {
        const spans = contentOzon('.mr2')

        spans.each((i, elem) => {
            if(contentOzon(elem).text().indexOf('шт.') >= 0) {
                // if(new_orders.indexOf(content(elem.parentNode.nextSibling).text().trim()) >= 0) {
                //     quantity.push(parseInt((content(elem).text().replace(' шт.', ''))))
                // }
                if(new_orders.indexOf(contentOzon(elem.parentNode.nextSibling).text().trim()) >= 0) {
                    quantity.push(parseInt((contentOzon(elem).text().replace(' шт.', ''))))
                }
            }
        })
        // spans.each((i, elem) => {
        //     if(content(elem).text().indexOf('шт.') === 0) quantity.push(content(elem).text())
        // })
    }

    getOrdersList(1,1)

    // console.log(new_orders.length)

    getQuantity()

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    const spans = content('span')

    const divs = content('.dDfDKJ')

    spans.each((i, elem) => {
        if(((content(elem).text()).indexOf('Полотенце') >= 0 || (content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
            products.push(content(elem).text())
        }
    })

    for(let i = 0; i < products.length; i++) {
        if(i%2 !== 0) {
            moderation_products.push(products[i])
        }
    }

    divs.each((i, elem) => {
        if((content(elem).text()).indexOf('029') >= 0) {
            moderation_gtins.push(content(elem).text())
        }
    })

    // console.log(moderation_products)
    // console.log(moderation_products.length)
    // console.log(moderation_gtins)
    // console.log(moderation_gtins.length)

    await wb.xlsx.readFile(hsFile)

    const ws = wb.getWorksheet('Краткий отчет')

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    html += `<div class="new_items_order">
                <h3>Товары не в каталоге</h3><hr>`

    new_orders.forEach(elem => {
        if(nat_cat.indexOf(elem) < 0) {
            let index = new_orders.indexOf(elem)
            html += `<p class="new">${elem} - <span>${quantity[index]} шт.</span></p>`
            new_items.push(elem)
            new_quantity.push(quantity[index])
        }
    })

    html += `<hr></div>
                    </section>${footerComponent}`

    function createNameList() {

                        let orderList = []
                        let _temp = []
                
                        for (let i = 0; i < new_items.length; i++) {
                
                            _temp.push(new_items[i])
                            
                                if(_temp.length%10 === 0) {
                                    orderList.push(_temp)
                                    _temp = []
                                }
                        }        
                
                        orderList.push(_temp)
                        _temp = []
                
                        return orderList
                
                    }
                
    function createQuantityList() {
                
                        let quantityList = []
                        let temp = []
                
                        for(let i = 0; i < new_quantity.length; i++) {
                
                            temp.push(new_quantity[i])
                
                                if(temp.length%10 === 0) {
                                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                                    temp = []
                                }
                
                        }
                
                        quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                
                        return quantityList
                
    }
                
    function createOrder() {
                
                let List = createNameList()
                let Quantity = createQuantityList()
                let content = ``
                
                for(let i = 0; i < List.length; i++) {
                    content += `<?xml version="1.0" encoding="utf-8"?>
                                    <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                        <lp>
                                            <productGroup>lp</productGroup>
                                            <contactPerson>333</contactPerson>
                                            <releaseMethodType>REMARK</releaseMethodType>
                                            <createMethodType>SELF_MADE</createMethodType>
                                            <productionOrderId>OZON</productionOrderId>
                                            <products>`
                for(let j = 0; j < List[i].length; j++) {
                    if(nat_cat.indexOf(List[i][j]) < 0) {
                    content += `<product>
                                    <gtin>${moderation_gtins[moderation_products.indexOf(List[i][j])]}</gtin>
                                    <quantity>${Quantity[i][j]}</quantity>
                                    <serialNumberType>OPERATOR</serialNumberType>
                                    <cisType>UNIT</cisType>
                                    <templateId>10</templateId>
                                </product>`
                }
            }
                
            content += `    </products>
                        </lp>
                    </order>`
                
            const date_ob = new Date()
                
            let month = date_ob.getMonth() + 1
                
            let filePath = ''
                
            month < 10 ? filePath = `./public/orders/lp_ozon_new_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_ozon_new_${i}_${date_ob.getDate()}_${month}.xml`
                
            fs.writeFileSync(filePath, content)
                
            content = ``
                
        }   
                
    }
                
    createOrder()

    res.send(html)

})

app.get('/wildberries', async function(req, res){
    
    const new_items = []
    const current_items = []
    const wb_orders = []
    const nat_cat = []
    const vendors = []
    const names = []
    const ozon = []

    let html = `${headerComponent}
                    <title>Импорт - WILDBERRIES</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

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

    // console.log(wb_orders)    

    wb_orders.forEach(elem => {
        if(vendors.indexOf(elem) >= 0){
            let index = vendors.indexOf(elem)
            // html += `<p>${names[index]}</p>`
            ozon.push(names[index])
            // console.log(typeof index)
        }
    })

    const testArray = []

    const test_Array = []

    ozon.forEach(elem => {
        if(testArray.indexOf(elem) < 0) {
            testArray.push(elem)
        }
    })

    for(let i = 0; i < testArray.length; i++) {
        let count = 0
        ozon.forEach(el => {
            if(testArray[i] === el) {
                count++
            }
        })
        test_Array.push(count)
    }

    // console.log(test_Array.length)
    // console.log(test_Array)    

    testArray.forEach(elem => {
        if(nat_cat.indexOf(elem) < 0 && new_items.indexOf(elem) < 0) {
            new_items.push(elem)            
        }

        if(nat_cat.indexOf(elem) >= 0 && current_items.indexOf(elem) < 0) {
            current_items.push(elem)
        }
    })

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    current_items.forEach(elem => {
        html += `<div class="table-row">
                    <span id="name">${elem}</span>
                    <span id="status-current">Актуальный товар</span>
                 </div>`
    })

    new_items.forEach(elem => {
        html += `<div class="table-row">
                    <span id="name">${elem}</span>
                    <span id="status-new">Новый товар</span>
                 </div>`
    })

    html += `</section>
             <section class="action-form">
                <button id="current-order"><a href="http://localhost:3030/wildberries_marks_order" target="_blank">Создать заказ маркировки для актуальных товаров</a></button>
                <button id="new-order"><a href="http://localhost:3030/wildberries_new_marks_order" target="_blank">Создать заказ маркировки для новых товаров</a></button>
             </section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

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
                    if(array[i].indexOf('Евро -') >= 0 || array[i].indexOf('евро -') >= 0 || array[i].indexOf('Евро на') >= 0 || array[i].indexOf('евро на') >= 0) {
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

        month < 10 ? filePath = `./public/wildberries/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_wildberries` : filePath = `./public/wildberries/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_wildberries`

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

    createImport(new_items)

    res.send(html)

})

app.get('/wildberries_marks_order', async function(req, res) {

    const wb_orders = []
    const nat_cat = []
    const vendors = []
    const names = []
    const ozon = []
    const gtins = []
    const orders = []

    const wb = new exl.Workbook()

    const hsFile = './public/Краткий отчет.xlsx'
    const ozonFile = './public/products.xlsx'
    const wbFile = './public/wildberries/new.xlsx'

    let html = `${headerComponent}
                    <title>Заказ актуальных маркировок</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>
             <section class="order-main">`

    await wb.xlsx.readFile(hsFile)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c_1 = ws.getColumn(1)

    c_1.eachCell(c => {
        gtins.push(c.value)        
    })

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

    // console.log(wb_orders)

    wb_orders.forEach(elem => {
        if(vendors.indexOf(elem) >= 0){
            let index = vendors.indexOf(elem)
            ozon.push(names[index])
            // console.log(typeof index)
        }
    })

    const testArray = []
    const test_Array = []

    ozon.forEach(elem => {
        if(testArray.indexOf(elem) < 0) {
            testArray.push(elem)            
        }
    })

    for(let i = 0; i < testArray.length; i++) {
        let count = 0
        let str = ''        
        ozon.forEach(el => {
            if(testArray[i] === el) {
                count++
            }
        })
        
        test_Array.push(count)
    }

    ozon.forEach(el => {
        if(orders.indexOf(el) < 0) orders.push(el)
    })

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < orders.length; i++) {

            _temp.push(orders[i])
            
                if(_temp.length%10 === 0) {
                    orderList.push(_temp)
                    _temp = []
                }
        }        

        orderList.push(_temp)
        _temp = []

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < test_Array.length; i++) {

            temp.push(test_Array[i])

                if(temp.length%10 === 0) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                    temp = []
                }

        }

        quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        return quantityList

    }

    function createOrder() {

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        for(let i = 0; i < List.length; i++) {
             content += `<?xml version="1.0" encoding="utf-8"?>
                        <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                            <lp>
                                <productGroup>lp</productGroup>
                                <contactPerson>333</contactPerson>
                                <releaseMethodType>REMARK</releaseMethodType>
                                <createMethodType>SELF_MADE</createMethodType>
                                <productionOrderId>WB</productionOrderId>
                                <products>`
            for(let j = 0; j < List[i].length; j++) {                
                if(nat_cat.indexOf(List[i][j]) >= 0) {
                    content += `<product>
                                    <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                    <quantity>${Quantity[i][j]}</quantity>
                                    <serialNumberType>OPERATOR</serialNumberType>
                                    <cisType>UNIT</cisType>
                                    <templateId>10</templateId>
                                </product>`
                }
            }

            content += `    </products>
                        </lp>
                    </order>`

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_${month}.xml`

            fs.writeFileSync(filePath, content)

            content = ``

        }
        
        // console.log(List)
        // console.log(Quantity)

        html += `<div class="new_items">
                    <h3>Список заказов</h3>
                        <hr>`

        for(let i = 0; i < List.length; i++) {
            for(let j = 0; j < List[i].length; j++) {
                if(nat_cat.indexOf(List[i][j]) < 0) {
                    html += `<p class="new">${List[i][j]} - <span>${Quantity[i][j]} шт.</span></p>`
                } else {
                    html += `<p class="current">${List[i][j]} - <span>${Quantity[i][j]} шт.</span></p>`
                }
            }
        }

        html += `<hr>
                    </div>                    
                        <section>${footerComponent}`

        

    }

    createOrder()

    res.send(html)

})

app.get('/wildberries_new_marks_order', async function(req, res){

    const wb = new exl.Workbook()

    const wb_orders = []
    const nat_cat = []
    const vendors = []
    const names = []
    const ozon = []
    const products = []
    const moderation_products = []
    const moderation_gtins = []
    const orders = []

    let html = `${headerComponent}
                    <title>Заказ новых маркировок</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>
             <section class="order-main">`

    const hsFile = './public/Краткий отчет.xlsx'
    const filePath = './public/moderation_marks/moderation_marks.html'
    const ozonFile = './public/products.xlsx'
    const wbFile = './public/wildberries/new.xlsx'

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    const spans = content('span')

    const divs = content('.dDfDKJ')

    spans.each((i, elem) => {
        if(((content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
            products.push(content(elem).text())
        }
    })

    for(let i = 0; i < products.length; i++) {
        if(i%2 !== 0) {
            moderation_products.push(products[i])
        }
    }

    divs.each((i, elem) => {
        if((content(elem).text()).indexOf('029') >= 0) {
            moderation_gtins.push(content(elem).text())
        }
    })

    // console.log(moderation_products)
    // console.log(moderation_products.length)
    // console.log(moderation_gtins)
    // console.log(moderation_gtins.length)

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

    // console.log(wb_orders)

    wb_orders.forEach(elem => {
        if(vendors.indexOf(elem) >= 0){
            let index = vendors.indexOf(elem)
            ozon.push(names[index])
            // console.log(typeof index)
        }
    })

    const testArray = []
    const test_Array = []

    ozon.forEach(elem => {
        if(testArray.indexOf(elem) < 0) {
            testArray.push(elem)            
        }
    })

    for(let i = 0; i < testArray.length; i++) {
        let count = 0
        let str = ''        
        ozon.forEach(el => {
            if(testArray[i] === el) {
                count++
            }
        })
        
        test_Array.push(count)
    }

    ozon.forEach(el => {
        if(orders.indexOf(el) < 0) orders.push(el)
    })

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < orders.length; i++) {

            
            _temp.push(orders[i])
            
            if(_temp.length === 10) {
                orderList.push(_temp)
                _temp = []
            }
            
        }        

        orderList.push(_temp)
        _temp = []

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < test_Array.length; i++) {

            temp.push(test_Array[i])

                if(temp.length === 10) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                    temp = []
                }

        }

        quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        return quantityList

    }

    function createOrder() {

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        html += `<div class="new_items">
                    <h3>Список заказов</h3>
                        <hr>`

        for(let i = 0; i < List.length; i++) {
            content += `<?xml version="1.0" encoding="utf-8"?>
                        <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                            <lp>
                                <productGroup>lp</productGroup>
                                <contactPerson>333</contactPerson>
                                <releaseMethodType>REMARK</releaseMethodType>
                                <createMethodType>SELF_MADE</createMethodType>
                                <productionOrderId>WB</productionOrderId>
                                <products>`
            for(let j = 0; j < List[i].length; j++) {
                if(nat_cat.indexOf(List[i][j]) < 0) {
                    html += `<p class="new">${List[i][j]} - <span>${Quantity[i][j]} шт.</span></p>`
                    content += `<product>
                                    <gtin>${moderation_gtins[moderation_products.indexOf(List[i][j])]}</gtin>
                                    <quantity>${Quantity[i][j]}</quantity>
                                    <serialNumberType>OPERATOR</serialNumberType>
                                    <cisType>UNIT</cisType>
                                    <templateId>10</templateId>
                                </product>`
                }
            }

            content += `    </products>
                        </lp>
                    </order>`

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_wb_new_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_wb_new_${i}_${date_ob.getDate()}_${month}.xml`

            fs.writeFileSync(filePath, content)

            content = ``

        }

        html += `<hr></div>
                    <section>${footerComponent}`

    }

    createOrder()

    res.send(html)

})

app.get('/input_remarking', async function(req, res){

    let html = `${headerComponent}
                    <title>Перемаркировка</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
                        
    let buttons = ['ozon', 'wb']
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {
            array[i] === 'wb' ? address = 'wildberries' : address = array[i]
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                     </button>`
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

    let remark_date = ''

    const date_ob = new Date()

    let year = date_ob.getFullYear()

    let month = date_ob.getMonth()+1

    let day = date_ob.getDate()

    month < 10 ? month = '0' + month : month

    day < 10 ? day = '0' + day : day

    remark_date = year + '-' + month + '-' + day    

    let content = `<?xml version="1.0" encoding="UTF-8"?>
                    <remark version="7">
                        <trade_participant_inn>372900043349</trade_participant_inn>
                        <remark_date>${remark_date}</remark_date>
                        <remark_cause>KM_SPOILED</remark_cause>
                            <products_list>`    

    const marks = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/inputinsale/marks.xlsx')

    const ws = wb.getWorksheet(1)

    ws.getColumn(1).eachCell(el => {
        marks.push(el.value.trim())
    })

    marks.forEach(el => {
        if(el.length === 31) {
            content += `<product>
                            <new_ki><![CDATA[${el}]]></new_ki>
                            <tnved_code_10>6302100001</tnved_code_10>
                            <production_country>РОССИЯ</production_country>
                        </product>`
        }
    })

    // console.log(content)   

    content += `    </products_list>
            </remark>`

    fs.writeFileSync('./public/inputinsale/remarking.xml', content)

    html += footerComponent

    res.send(html)
    
})

app.get('/sale_ozon', async function(req, res){

    const date_ob = new Date()

    let date_string = ''

    let [year, month, day] = [date_ob.getFullYear(), date_ob.getMonth()+1, date_ob.getDate()]

    month < 10 ? month = '0' + month : month
    day < 10 ? day = '0' + day : day

    date_string = `${year}-${month}-${day}`

    let content = `<?xml version="1.0" encoding="utf-8"?>
                    <withdrawal version="8">
                        <trade_participant_inn>372900043349</trade_participant_inn>
                        <withdrawal_type>DISTANCE</withdrawal_type>
                        <withdrawal_date>${date_string}</withdrawal_date>
                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                        <products_list>`
    
    const wb = new exl.Workbook()

    async function getNationalCatalog() {

        const filePath = './public/Краткий отчет.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('Краткий отчет')

        const c2 = ws.getColumn(2)

        const gtins = []
        const names = []

        const c1 = ws.getColumn(1)

        c1.eachCell(c => {
            gtins.push(c.value)
        })

        c2.eachCell(c => {
            // console.log(c.value)
            names.push(c.value)
        })

        return [names, gtins]

    }

    //получаем содержимое файла нац. каталога
    //а именно - наименование уже созданные ранее
    //помещаем их в отдельный массив

    const [catalogNames, catalogGtins] = await getNationalCatalog()

    // console.log(catalogNames)

    //получаем данные из xlsx файла с реализациями и
    //формируем массив объектов реализаций

    async function getConsignments() {

        const consignmentNumbers = []

        const consignmentProducts = []

        const noRepeatConsignmentNumbers = []

        const consignmentTypes = []

        const noRepeatConsignmentTypes = []

        const consignments = []

        const filePath = './public/distance/релизации.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('TDSheet')

        const [c2, c4, c6] = [ws.getColumn(2), ws.getColumn(4), ws.getColumn(6)]

        c2.eachCell(c => {
            let str = c.value
            consignmentNumbers.push(str.substring(str.length - 4))
        })

        c4.eachCell(c => {
            consignmentProducts.push(c.value)
        })

        c6.eachCell(c => {
            consignmentTypes.push(c.value)
        })
        

        for(let i = 0; i < consignmentNumbers.length; i++) {

            if(consignmentTypes[i] !== null && consignmentTypes[i].indexOf('ozon') >= 0 && noRepeatConsignmentNumbers.indexOf(consignmentNumbers[i]) < 0 && catalogNames.indexOf(consignmentProducts[i]) >= 0) {

                noRepeatConsignmentNumbers.push(consignmentNumbers[i])

            }

        }

        // console.log(consignmentTypes)
        // console.log(noRepeatConsignmentNumbers)

        for(let i = 0; i < consignmentTypes.length; i++) {

            if(consignmentTypes[i] !== null && consignmentTypes[i].indexOf('ozon') >= 0 && catalogNames.indexOf(consignmentProducts[i]) >= 0) {

                let str = consignmentTypes[i]
                if(noRepeatConsignmentTypes.indexOf(str.substring(5)) < 0) {

                    noRepeatConsignmentTypes.push(str.substring(5))

                }

            }

        }
        
        for(let i = 0; i < noRepeatConsignmentNumbers.length; i++) {

            let elem = {
                number: noRepeatConsignmentNumbers[i],
                products: [],
                ozonNumber: ''
            }

            for(let i = 0; i < consignmentProducts.length; i++) {
                if(consignmentNumbers[i] === elem.number) {
                    elem.products.push(consignmentProducts[i])
                }
            }

            let index = consignmentNumbers.indexOf(noRepeatConsignmentNumbers[i])
            let str = consignmentTypes[index]
            elem.ozonNumber = str.substring(5)

            consignments.push(elem)

        }

        return consignments

    }

    const consNumbers = await getConsignments()

    console.log(consNumbers)

    //аналогичный предыдущему метод
    //для формирования массива объектов
    //заказов Ozon

    async function getOzonOrders() {

        const filePath = './public/distance/postings.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('Worksheet')

        const orderNumbers = []
        const orderProducts = []
        const orderQuantitys = []
        const orderCosts = []
        const orders = []

        const [c2, c9, c12, c14] = [ws.getColumn(2), ws.getColumn(9), ws.getColumn(12), ws.getColumn(14)]

        c2.eachCell(c => {
            orderNumbers.push(c.value)
        })

        c9.eachCell(c => {
            orderProducts.push(c.value)
        })

        c14.eachCell(c => {
            orderQuantitys.push(c.value)
        })

        c12.eachCell(c => {
            orderCosts.push(c.value)
        })

        const noRepeatOrderNumbers = []

        for(let i = 0; i < orderNumbers.length; i++) {

            if(noRepeatOrderNumbers.indexOf(orderNumbers[i]) < 0 && catalogNames.indexOf(orderProducts[i]) >= 0) {
                noRepeatOrderNumbers.push(orderNumbers[i])
            }

        }

        // console.log(noRepeatOrderNumbers)
        // console.log(noRepeatOrderNumbers.length)

        for(let j = 0; j < noRepeatOrderNumbers.length; j++) {
            // console.log('+')
            let elem = {
                number: noRepeatOrderNumbers[j],
                products: [],
                quantitys: [],
                costs: []
            }

            for(let j = 0; j < orderProducts.length; j++) {

                if(orderNumbers[j] == elem.number) {
                    elem.products.push(orderProducts[j])
                    elem.quantitys.push(orderQuantitys[j])
                    elem.costs.push(orderCosts[j])
                }

            }           

            orders.push(elem)

        }

        return orders

    }

    const orderNumbers = await getOzonOrders()

    console.log(orderNumbers)

    async function getMarks() {

        const filePath = './public/distance/marks.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('Worksheet')

        const allMarks = []
        const allGtins = []
        const marks = []

        const c1 = ws.getColumn(1)

        const c2 = ws.getColumn(2)

        c1.eachCell(c => {
            allMarks.push(c.value)
        })

        c2.eachCell(c => {
            if(c.value !== null) {
            let str = c.value
            allGtins.push(str.substring(1))
            } else {
                allGtins.push(c.value)
            }
        })

        const noRepeatGtins = []

        for(let i = 0; i < allGtins.length; i++) {
            if(noRepeatGtins.indexOf(allGtins[i]) < 0) {
                noRepeatGtins.push(allGtins[i])
            }
        }

        for(let i = 0; i < noRepeatGtins.length; i++) {

            let elem = {
                gtin: noRepeatGtins[i],
                product:'',
                marks: [],
                quantity: 0
            }

            for(let i = 0; i < allMarks.length; i++) {
                if(allGtins[i] == elem.gtin) {
                    elem.marks.push(allMarks[i])
                    elem.product = catalogNames[catalogGtins.indexOf(elem.gtin)]
                    elem.quantity++
                }
            }

            marks.push(elem)

        }

        return marks

    }

    const introducedMarks = await getMarks()

    console.log(introducedMarks)

    //Создаем цикл для формирования списка товаров
    //Подлежащих выводу из оборота

    async function createSaleDocument() {

        for(let j = 0; j < consNumbers.length; j++) {
            if(consNumbers[j].ozonNumber !== undefined && consNumbers[j].ozonNumber !== null) {
                for(let i = 0; i < orderNumbers.length; i++) {
                    
                    if(orderNumbers[i].number == consNumbers[j].ozonNumber) {
                        for(let k = 0; k < orderNumbers[i].products.length; k++) {
    
                            if(catalogNames.indexOf(orderNumbers[i].products[k]) >= 0) {
                                
                                let index = introducedMarks.findIndex(el => el.gtin === catalogGtins[catalogNames.indexOf(orderNumbers[i].products[k])])
                                if(index >= 0) {
                                    if(introducedMarks[index].quantity == orderNumbers[i].quantitys[k]) {
                                        introducedMarks[index].marks.forEach(el => {
                                            // console.log(el)
                                            content += `<product>
                                                            <cis><![CDATA[${el}]]></cis>
                                                            <cost>${orderNumbers[i].costs[k]}00</cost>
                                                            <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                                                            <primary_document_number>${consNumbers[j].number}</primary_document_number>
                                                            <primary_document_date>2023-05-22</primary_document_date>
                                                        </product>`
                                        })
                                    } else {    
                                        content += `<product>
                                                        <cis><![CDATA[${introducedMarks[index].marks[0]}]]></cis>
                                                        <cost>${orderNumbers[i].costs[k]}00</cost>
                                                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                                                        <primary_document_number>${consNumbers[j].number}</primary_document_number>
                                                        <primary_document_date>2023-05-22</primary_document_date>
                                                    </product>`    
                                    }
                                }
    
                            }
    
                        }
                        
                    }
                    
                }
            }
        }    

        content += `</products_list>
                </withdrawal>`
    
        const filePath = `./public/distance/ozon_distance_${date_string}.xml`
    
        fs.writeFileSync(filePath, content)

    }

    await createSaleDocument()

    res.send(`It's working...`)

})

app.get('/sale_wb', async function(req, res){

    const date_ob = new Date()

    let date_string = ''

    let [year, month, day] = [date_ob.getFullYear(), date_ob.getMonth()+1, date_ob.getDate()]

    month < 10 ? month = '0' + month : month
    day < 10 ? day = '0' + day : day

    date_string = `${year}-${month}-${day}`

    let content = `<?xml version="1.0" encoding="utf-8"?>
                    <withdrawal version="8">
                        <trade_participant_inn>372900043349</trade_participant_inn>
                        <withdrawal_type>DISTANCE</withdrawal_type>
                        <withdrawal_date>${date_string}</withdrawal_date>
                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                        <products_list>`
    
    const wb = new exl.Workbook()

    async function getNationalCatalog() {

        const filePath = './public/Краткий отчет.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('Краткий отчет')

        const c2 = ws.getColumn(2)

        const gtins = []
        const names = []

        const c1 = ws.getColumn(1)

        c1.eachCell(c => {
            gtins.push(c.value)
        })

        c2.eachCell(c => {
            // console.log(c.value)
            names.push(c.value)
        })

        return [names, gtins]

    }

    //получаем содержимое файла нац. каталога
    //а именно - наименование уже созданные ранее
    //помещаем их в отдельный массив

    const [catalogNames, catalogGtins] = await getNationalCatalog()

    async function getConsignments() {

        const consignmentNumbers = []

        const consignmentProducts = []

        const noRepeatConsignmentNumbers = []

        const consignmentTypes = []

        const noRepeatConsignmentTypes = []

        const consignments = []

        const filePath = './public/distance/релизации.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('TDSheet')

        const [c2, c4, c6] = [ws.getColumn(2), ws.getColumn(4), ws.getColumn(6)]

        c2.eachCell(c => {
            let str = c.value
            consignmentNumbers.push(str.substring(str.length - 4))        
        })

        c4.eachCell(c => {
            consignmentProducts.push(c.value)
        })

        c6.eachCell(c => {
            consignmentTypes.push(c.value)
        })
        

        for(let i = 0; i < consignmentNumbers.length; i++) {

            if(consignmentTypes[i] !== null && consignmentTypes[i].indexOf('WB') >= 0 && noRepeatConsignmentNumbers.indexOf(consignmentNumbers[i]) < 0 && catalogNames.indexOf(consignmentProducts[i]) >= 0) {

                noRepeatConsignmentNumbers.push(consignmentNumbers[i])

            }

        }

        // console.log(consignmentTypes)

        for(let i = 0; i < consignmentTypes.length; i++) {

            if(consignmentTypes[i] !== null && consignmentTypes[i].indexOf('WB') >= 0 && catalogNames.indexOf(consignmentProducts[i]) >= 0) {

                let str = consignmentTypes[i]
                if(noRepeatConsignmentTypes.indexOf(str.substring(5)) < 0) {

                    noRepeatConsignmentTypes.push(str.substring(5))

                }

            }

        }
        
        for(let i = 0; i < noRepeatConsignmentNumbers.length; i++) {

            let elem = {
                number: noRepeatConsignmentNumbers[i],
                products: [],
                wbNumber: ''
            }

            for(let i = 0; i < consignmentProducts.length; i++) {
                if(consignmentNumbers[i] === elem.number) {
                    elem.products.push(consignmentProducts[i])
                }
            }

            let index = consignmentNumbers.indexOf(noRepeatConsignmentNumbers[i])
            let str = consignmentTypes[index]
            elem.wbNumber = str.substring(3)

            consignments.push(elem)

        }

        return consignments

    }

    const consNumbers = await getConsignments()

    // console.log(consNumbers)

    async function getWBOrders() {

        const filePath = './public/distance/wb_orders.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('КИЗ,УИН')

        const [c3, c5] = [ws.getColumn(3), ws.getColumn(5)]

        const [orderCis, orderCost] = [[], []]

        c3.eachCell(c => {
            orderCis.push(c.value)
        })

        c5.eachCell(c => {
            orderCost.push(c.value)
        })

        return [orderCis, orderCost]

    }

    const [ordersCis, ordersCost] = await getWBOrders()

    async function getMarks() {

        const filePath = './public/distance/marks.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('Worksheet')

        const allMarks = []
        const allGtins = []
        const marks = []

        const c1 = ws.getColumn(1)

        const c2 = ws.getColumn(2)

        c1.eachCell(c => {
            allMarks.push(c.value)
        })

        c2.eachCell(c => {
            if(c.value !== null) {
            let str = c.value
            allGtins.push(str.substring(1))
            } else {
                allGtins.push(c.value)
            }
        })

        const noRepeatGtins = []

        for(let i = 0; i < allGtins.length; i++) {
            if(noRepeatGtins.indexOf(allGtins[i]) < 0) {
                noRepeatGtins.push(allGtins[i])
            }
        }

        for(let i = 0; i < noRepeatGtins.length; i++) {

            let elem = {
                gtin: noRepeatGtins[i],
                product:'',
                marks: [],
                quantity: 0
            }

            for(let i = 0; i < allMarks.length; i++) {
                if(allGtins[i] == elem.gtin) {
                    elem.marks.push(allMarks[i])
                    elem.product = catalogNames[catalogGtins.indexOf(elem.gtin)]
                    elem.quantity++
                }
            }

            marks.push(elem)

        }

        return marks

    }

    const introducedMarks = await getMarks()

    //Создаем цикл для формирования списка товаров
    //Подлежащих выводу из оборота

    async function createSaleDocument() {

        for(let i = 0; i < ordersCis.length; i++) {
            let number = ''
            let index = introducedMarks.findIndex(el => el.marks == ordersCis[i])
            let idx = null
            if(index >= 0) {
                idx = consNumbers.findIndex(el => el.products == introducedMarks[index].product)

                if(idx >= 0) {
                    content += `<product>
                            <cis><![CDATA[${ordersCis[i]}]]></cis>
                            <cost>${ordersCost[i]}00</cost>
                            <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                            <primary_document_number>${consNumbers[idx].number}</primary_document_number>
                            <primary_document_date>2023-05-22</primary_document_date>
                        </product>`
                }

                

            }

            

        }
                            
        content += `</products_list>
                </withdrawal>`
    
        const filePath = `./public/distance/wb_distance_${date_string}.xml`
    
        fs.writeFileSync(filePath, content)

    }

    await createSaleDocument()

    res.send(`It's working...`)

})

app.listen(3030)