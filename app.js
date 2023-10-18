const express = require('express')
const exl = require('exceljs')
const fs = require('fs')
const cio = require('cheerio')
const fetch = require('node-fetch')
const { connect } = require('http2')
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
                            <p class="nav-item" id="home"><a href="http://localhost:3030/home">Главная</a></p>
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

let buttons = ['ozon', 'wb', 'sber', 'yandex']

app.use(express.static(__dirname + '/public'))

app.get('/home', async function(req, res){

    let html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
        }

        html += `   </div>`

    }

    async function renderMarkingButtons() {
        html += `<div class="marking-control">
                    <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота OZON (Дистанционная продажа)</a></button>
                    <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота WILDBERRIES (Дистанционная продажа)</a></button>
                 </div>`
    }

    await renderImportButtons(buttons)
    await renderMarkingButtons()

    html += `</section>`

    html += `<section class="filter-control">
                <div class="search-field">
                    <input class="search-input" type="text" placeholder="Код или GTIN товара">
                    <button id="search" type="submit">
                        <svg width="20" height="20" fill="none" xmlns="http://www.w3.org/2000/svg" cursor="default" style="color: rgb(122, 129, 155);"><path fill-rule="evenodd" clip-rule="evenodd" d="M10.75 1.739a.75.75 0 00-1.5 0V9.25H1.739a.75.75 0 100 1.5H9.25V18.261H10h-.75a.75.75 0 101.5 0H10h.75V10.75H18.261V10v.75a.75.75 0 000-1.5V10v-.75H10.75V1.739z" fill="currentColor">
                        </path></svg>
                    </button>
                </div>
                <div class="multiple-list">
                    <div class="multiple-status">
                        Статус
                    </div>
                    <div class="status-list">
                        <ul class="list">
                            <li class="list-item">Нанесен</li>
                            <li class="list-item">В обороте</li>
                            <li class="list-item">Выбыл</li>
                        </ul>
                    </div>
                    <svg width="16" height="16" fill="none" xmlns="http://www.w3.org/2000/svg" class="MuiSelect-icon MuiSelect-iconStandard css-1rb0eps"><path d="M12 6H4l4 4 4-4z" fill="currentColor">
                    </path></svg>
                </div>
                <button class="show-button"><a id="show-anchor">Показать</a></button>
             </section>`

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
        const actual_status = []

        const wb = new exl.Workbook()

        await wb.xlsx.readFile('./public/actual_marks.xlsx')

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c2, c16, c23] = [ws.getColumn(1), ws.getColumn(2), ws.getColumn(16), ws.getColumn(23)]

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

        c16.eachCell(c => {
            if(c.value != null && c.value != 'status') {
                actual_status.push(c.value)
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

        return [actual_gtins, actual_marks, actual_dates, actual_status]

    }

    async function renderMarksTable() {
        
        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

        async function createPages(array) {

            let marks_list = []
            let _temp = []

            for(let i = 0; i < array.length; i++) {

                _temp.push({
                    gtin: actual_gtins[i],
                    mark: array[i],
                    date: actual_dates[i],
                    status: actual_status[i],
                    order: ''
                })

                if(_temp.length%10 === 0) {
                    marks_list.push(_temp)
                    _temp = []
                }

            }

            marks_list.push(_temp)
            _temp = []

            return marks_list

        }

        let pageNumber = 0

        if(req.query.page == null || req.query.page == undefined || req.query.page == 0) {

            pageNumber = 1

        } else {

            pageNumber = parseInt(req.query.page)

        }

        // if(req.query.order == null || req.query.order == undefined || req.query.order == 0) {
            
        // } else {

        //     let page = 0
        //     let index = Pages[page].findIndex(el => el.mark == req.query.mark)
            
        //     Pages[page]

        // }

        let Pages = await createPages(actual_marks)

        html += `<section class="table">
                            <div class="marks-table">
                                <div class="marks-table-header">
                                    <div class="header-cell">КИЗ</div>
                                    <div class="header-cell">GTIN</div>
                                    <div class="header-cell">Товар</div>
                                    <div class="header-cell">Дата эмиссии</div>
                                    <div class="header-cell">Статус</div>
                                    <!--<div class="header-cell">Номер заказа</div>-->
                                </div>
                                <div class="header-wrapper"></div>`

        for(let j = 0; j < Pages[pageNumber - 1].length; j++) {

                let status = ''
                if(Pages[pageNumber - 1][j].status == 'INTRODUCED') {
                    status = 'В обороте'
                } else if(Pages[pageNumber - 1][j].status == 'APPLIED') {
                    status = 'Нанесен'
                } else if(Pages[pageNumber - 1][j].status == 'RETIRED') {
                    status = 'Выбыл'
                }
                    
                html+= `<div class="table-row">
                            <span type="text" id="mark">${Pages[pageNumber - 1][j].mark}</span>
                            <span id="gtin">${Pages[pageNumber - 1][j].gtin}</span>
                            <span id="name">${names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)] == undefined ? '-' : names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)]}</span>
                            <span id="date">${Pages[pageNumber - 1][j].date}</span>
                            <span id="status">${status}</span>
                            <!--<div>
                                <input id="order" type="text" placeholder="${Pages[pageNumber - 1][j].order}">
                                <button type="submit"><a class="order-number" href="">Отправить</a></button>
                            </div>-->
                        </div>`
                
            }
        
        return Math.ceil(Pages.length)
    
    }


    let lastPage = await renderMarksTable()

    html += `       </div>
                <div class="pages">
                    <a id="begin" href="http://localhost:3030/home">На первую страницу</a>
                    <div class="pages-prev">
                        <svg id="prev-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M4.113 9.669c.432.441 1.13.441 1.563 0a1.145 1.145 0 0 0 0-1.596L2.668 4.999l3.008-3.072a1.145 1.145 0 0 0 0-1.596 1.087 1.087 0 0 0-1.563 0l-3.79 3.87A1.14 1.14 0 0 0 0 5c0 .29.108.578.324.799l3.79 3.87z">
                        </path></svg>
                        <a id="prev" href="">Предыдущая страница</a>
                    </div>
                    <div class="pages-next">
                        <a id="next" href="">Следующая страница</a>
                        <svg id="next-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M1.887.331a1.087 1.087 0 0 0-1.563 0 1.145 1.145 0 0 0 0 1.596l3.008 3.074L.324 8.073a1.145 1.145 0 0 0 0 1.596c.432.441 1.13.441 1.563 0l3.79-3.87A1.14 1.14 0 0 0 6 5c0-.29-.108-.578-.324-.799L1.886.332z">
                        </path></svg>
                    </div>
                    <a id="last" class="pages-last" href="http://localhost:3030/home?page=${lastPage}">На последнюю страницу</a>                  
                </div>
            </section>
        <div class="body-wrapper"></div>
    ${footerComponent}`

    res.send(html)
})

app.get('/home/:status/', async function(req, res){

    let html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    html += `<section class="filter-control">
                <div class="search-field">
                    <input class="search-input" type="text" placeholder="Код или GTIN товара">
                    <button id="search" type="submit">
                        <svg width="20" height="20" fill="none" xmlns="http://www.w3.org/2000/svg" cursor="default" style="color: rgb(122, 129, 155);"><path fill-rule="evenodd" clip-rule="evenodd" d="M10.75 1.739a.75.75 0 00-1.5 0V9.25H1.739a.75.75 0 100 1.5H9.25V18.261H10h-.75a.75.75 0 101.5 0H10h.75V10.75H18.261V10v.75a.75.75 0 000-1.5V10v-.75H10.75V1.739z" fill="currentColor">
                        </path></svg>
                    </button>
                </div>
                <div class="multiple-list">
                    <div class="multiple-status">
                        Статус
                    </div>
                    <div class="status-list">
                        <ul class="list">
                            <li class="list-item">Нанесен</li>
                            <li class="list-item">В обороте</li>
                            <li class="list-item">Выбыл</li>
                        </ul>
                    </div>
                    <svg width="16" height="16" fill="none" xmlns="http://www.w3.org/2000/svg" class="MuiSelect-icon MuiSelect-iconStandard css-1rb0eps"><path d="M12 6H4l4 4 4-4z" fill="currentColor">
                    </path></svg>
                </div>
                <button class="show-button"><a id="show-anchor">Показать</a></button>
             </section>`

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
        const actual_status = []

        const wb = new exl.Workbook()

        await wb.xlsx.readFile('./public/actual_marks.xlsx')

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c2, c16, c23] = [ws.getColumn(1), ws.getColumn(2), ws.getColumn(16), ws.getColumn(23)]

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

        c16.eachCell(c => {
            if(c.value != null && c.value != 'status') {
                actual_status.push(c.value)
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

        return [actual_gtins, actual_marks, actual_dates, actual_status]

    }

    async function renderMarksTable() {

        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

        async function createFilterPages(array, status) {

            let marks_list = []
    
            let _temp = []
    
            if(status == 'APPLIED') {
    
                for(let i = 0; i < array.length; i++) {
    
                    if(status == actual_status[i]) {
    
                        _temp.push({
                            gtin: actual_gtins[i],
                            mark: array[i],
                            date: actual_dates[i],
                            status: actual_status[i],
                            order: ''
                        })
    
                        if(_temp.length%10 === 0) {
                            marks_list.push(_temp)
                            _temp = []
                        }
    
                    }
    
                }
    
                marks_list.push(_temp)
                _temp = []
    
            }
    
            if(status == 'RETIRED') {
    
                for(let i = 0; i < array.length; i++) {
    
                    if(status == actual_status[i]) {
    
                        _temp.push({
                            gtin: actual_gtins[i],
                            mark: array[i],
                            date: actual_dates[i],
                            status: actual_status[i],
                            order: ''
                        })
    
                        if(_temp.length%10 === 0) {
                            marks_list.push(_temp)
                            _temp = []
                        }
    
                    }
    
                }
    
                marks_list.push(_temp)
                _temp = []
            }
    
            if(status == 'INTRODUCED') {
    
                for(let i = 0; i < array.length; i++) {
    
                    if(status == actual_status[i]) {
    
                        _temp.push({
                            gtin: actual_gtins[i],
                            mark: array[i],
                            date: actual_dates[i],
                            status: actual_status[i],
                            order: ''
                        })
    
                        if(_temp.length%10 === 0) {
                            marks_list.push(_temp)
                            _temp = []
                        }
    
                    }
    
                }
    
                marks_list.push(_temp)
                _temp = []
    
            }
    
            return marks_list
    
        }

        let pageNumber = 0

        if(req.query.page == null || req.query.page == undefined || req.query.page == 0) {

            pageNumber = 1

        } else {

            pageNumber = parseInt(req.query.page)

        }

        // if(req.query.order == null || req.query.order == undefined || req.query.order == 0) {
            
        // } else {

        //     let page = 0
        //     let index = Pages[page].findIndex(el => el.mark == req.query.mark)
            
        //     Pages[page]

        // }

        let Pages = await createFilterPages(actual_marks, req.params.status)

        html += `<section class="table">
                            <div class="marks-table">
                                <div class="marks-table-header">
                                    <div class="header-cell">КИЗ</div>
                                    <div class="header-cell">GTIN</div>
                                    <div class="header-cell">Товар</div>
                                    <div class="header-cell">Дата эмиссии</div>
                                    <div class="header-cell">Статус</div>
                                    <!--<div class="header-cell">Номер заказа</div>-->
                                </div>
                                <div class="header-wrapper"></div>`

        for(let j = 0; j < Pages[pageNumber - 1].length; j++) {

                let status = ''
                if(Pages[pageNumber - 1][j].status == 'INTRODUCED') {
                    status = 'В обороте'
                } else if(Pages[pageNumber - 1][j].status == 'APPLIED') {
                    status = 'Нанесен'
                } else if(Pages[pageNumber - 1][j].status == 'RETIRED') {
                    status = 'Выбыл'
                }
                    
                html+= `<div class="table-row">
                            <span type="text" id="mark">${Pages[pageNumber - 1][j].mark}</span>
                            <span id="gtin">${Pages[pageNumber - 1][j].gtin}</span>
                            <span id="name">${names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)] == undefined ? '-' : names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)]}</span>
                            <span id="date">${Pages[pageNumber - 1][j].date}</span>
                            <span id="status">${status}</span>
                            <!--<div>
                                <input id="order" type="text" placeholder="${Pages[pageNumber - 1][j].order}">
                                <button type="submit"><a class="order-number" href="">Отправить</a></button>
                            </div>-->
                        </div>`
                
            }
        
        return Math.ceil(Pages.length)

    }

    let lastPage = await renderMarksTable()

    html += `       </div>
                <div class="pages">
                    <a id="begin" href="http://localhost:3030/home/${req.params.status}">На первую страницу</a>
                    <div class="pages-prev">
                        <svg id="prev-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M4.113 9.669c.432.441 1.13.441 1.563 0a1.145 1.145 0 0 0 0-1.596L2.668 4.999l3.008-3.072a1.145 1.145 0 0 0 0-1.596 1.087 1.087 0 0 0-1.563 0l-3.79 3.87A1.14 1.14 0 0 0 0 5c0 .29.108.578.324.799l3.79 3.87z">
                        </path></svg>
                        <a id="prev" href="">Предыдущая страница</a>
                    </div>
                    <div class="pages-next">
                        <a id="next" href="">Следующая страница</a>
                        <svg id="next-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M1.887.331a1.087 1.087 0 0 0-1.563 0 1.145 1.145 0 0 0 0 1.596l3.008 3.074L.324 8.073a1.145 1.145 0 0 0 0 1.596c.432.441 1.13.441 1.563 0l3.79-3.87A1.14 1.14 0 0 0 6 5c0-.29-.108-.578-.324-.799L1.886.332z">
                        </path></svg>
                    </div>
                    <a id="last" class="pages-last" href="http://localhost:3030/home/${req.params.status}?page=${lastPage}">На последнюю страницу</a>                  
                </div>
            </section>
        <div class="body-wrapper"></div>
    ${footerComponent}`

    res.send(html)

})

app.get('/filter', async function(req, res) {
    let html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    html += `<section class="filter-control">
                <div class="search-field">
                    <input class="search-input" type="text" placeholder="Код или GTIN товара">
                    <button id="search" type="submit">
                        <svg width="20" height="20" fill="none" xmlns="http://www.w3.org/2000/svg" cursor="default" style="color: rgb(122, 129, 155);"><path fill-rule="evenodd" clip-rule="evenodd" d="M10.75 1.739a.75.75 0 00-1.5 0V9.25H1.739a.75.75 0 100 1.5H9.25V18.261H10h-.75a.75.75 0 101.5 0H10h.75V10.75H18.261V10v.75a.75.75 0 000-1.5V10v-.75H10.75V1.739z" fill="currentColor">
                        </path></svg>
                    </button>
                </div>
                <div class="multiple-list">
                    <div class="multiple-status">
                        Статус
                    </div>
                    <div class="status-list">
                        <ul class="list">
                            <li class="list-item">Нанесен</li>
                            <li class="list-item">В обороте</li>
                            <li class="list-item">Выбыл</li>
                        </ul>
                    </div>
                    <svg width="16" height="16" fill="none" xmlns="http://www.w3.org/2000/svg" class="MuiSelect-icon MuiSelect-iconStandard css-1rb0eps"><path d="M12 6H4l4 4 4-4z" fill="currentColor">
                    </path></svg>
                </div>
                <button class="show-button"><a id="show-anchor">Показать</a></button>
             </section>`

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
        const actual_status = []

        const wb = new exl.Workbook()

        await wb.xlsx.readFile('./public/actual_marks.xlsx')

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c2, c16, c23] = [ws.getColumn(1), ws.getColumn(2), ws.getColumn(16), ws.getColumn(23)]

        c1.eachCell(c => {
            if(c.value.indexOf('01') >= 0) {
                let str = c.value
                if(str.indexOf('&') >= 0) {
                    str = str.replace(/&/g, '&amp;')
                }
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

        c16.eachCell(c => {
            if(c.value != null && c.value != 'status') {
                actual_status.push(c.value)
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

        return [actual_gtins, actual_marks, actual_dates, actual_status]

    }

    async function renderMarksTable() {

        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

        html += `<section class="table">
                            <div class="marks-table">
                                <div class="marks-table-header">
                                    <div class="header-cell">КИЗ</div>
                                    <div class="header-cell">GTIN</div>
                                    <div class="header-cell">Товар</div>
                                    <div class="header-cell">Дата эмиссии</div>
                                    <div class="header-cell">Статус</div>
                                    <!--<div class="header-cell">Номер заказа</div>-->
                                </div>
                                <div class="header-wrapper"></div>`

        if(req.query.cis != '' && req.query.gtin == undefined) {

            // console.log(req.query.cis)
            // let str = req.query.cis.replace(/</g, '&lt;')
            // console.log(str)
            // console.log(actual_marks[8])

            let index = 0
        
            for(let i = 0; i < actual_marks.length; i++) {

                if(actual_marks[i].indexOf(req.query.cis) >= 0) {

                    index = i

                }

            }
        
            let status = ''
                    if(actual_status[index] == 'INTRODUCED') {
                        status = 'В обороте'
                    } else if(actual_status[index] == 'APPLIED') {
                        status = 'Нанесен'
                    } else if(actual_status[index] == 'RETIRED') {
                        status = 'Выбыл'
                    }
                        
                    html+= `<div class="table-row">
                                <span type="text" id="mark">${actual_marks[index]}</span>
                                <span id="gtin">${actual_gtins[index]}</span>
                                <span id="name">${names[gtins.indexOf(actual_gtins[index])]}</span>
                                <span id="date">${actual_dates[index]}</span>
                                <span id="status">${status}</span>
                                <!--<div>
                                    <input id="order" type="text" placeholder="">
                                    <button type="submit"><a class="order-number" href="">Отправить</a></button>
                                </div>-->
                            </div>`
        }

        if(req.query.gtin != '' && req.query.cis == undefined) {

            for(let i = 0; i < actual_marks.length; i++) {

                if(actual_gtins[i] == req.query.gtin) {

                    let status = ''
                    if(actual_status[i] == 'INTRODUCED') {
                        status = 'В обороте'
                    } else if(actual_status[i] == 'APPLIED') {
                        status = 'Нанесен'
                    } else if(actual_status[i] == 'RETIRED') {
                        status = 'Выбыл'
                    }
                        
                    html+= `<div class="table-row">
                                <span type="text" id="mark">${actual_marks[i]}</span>
                                <span id="gtin">${actual_gtins[i]}</span>
                                <span id="name">${names[gtins.indexOf(actual_gtins[i])]}</span>
                                <span id="date">${actual_dates[i]}</span>
                                <span id="status">${status}</span>
                                <!--<div>
                                    <input id="order" type="text" placeholder="">
                                    <button type="submit"><a class="order-number" href="">Отправить</a></button>
                                </div>-->
                            </div>`

                }

            }

        }

    }

    await renderMarksTable()

    html += `</section>
        <div class="body-wrapper"></div>
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
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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
    const moderation_items = []
    const vendorCodes = []

    const colors = ['БЕЖЕВЫЙ', 'БЕЛЫЙ', 'БИРЮЗОВЫЙ', 'БОРДОВЫЙ', 'БРОНЗОВЫЙ', 'ВАНИЛЬ', 'ВИШНЯ', 'ГОЛУБОЙ', 'ЖЁЛТЫЙ', 'ЗЕЛЁНЫЙ', 'ЗОЛОТОЙ', 'ИЗУМРУДНЫЙ',
    'КАПУЧИНО', 'КИРПИЧНЫЙ', 'КОРАЛЛОВЫЙ', 'КОРИЧНЕВЫЙ', 'КРАСНЫЙ', 'ЛАЙМ', 'ЛЕОПАРД', 'МАЛИНОВЫЙ', 'МЕДНЫЙ', 'МОЛОЧНЫЙ', 'МЯТНЫЙ', 'ОЛИВКОВЫЙ', 'ОРАНЖЕВЫЙ',
    'ПЕСОЧНЫЙ', 'ПЕРСИКОВЫЙ', 'ПУРПУРНЫЙ', 'РАЗНОЦВЕТНЫЙ', 'РОЗОВО-БЕЖЕВЫЙ', 'РОЗОВЫЙ', 'СВЕТЛО-БЕЖЕВЫЙ', 'СВЕТЛО-ЗЕЛЕНЫЙ', 'СВЕТЛО-КОРИЧНЕВЫЙ', 'СВЕТЛО-РОЗОВЫЙ',
    'СВЕТЛО-СЕРЫЙ', 'СВЕТЛО-СИНИЙ', 'СВЕТЛО-ФИОЛЕТОВЫЙ', 'СЕРЕБРЯНЫЙ', 'СЕРО-ЖЕЛТЫЙ', 'СЕРО-ГОЛУБОЙ', 'СЕРЫЙ', 'СИНИЙ', 'СИРЕНЕВЫЙ', 'ЛИЛОВЫЙ', 'СЛИВОВЫЙ',
    'ТЕМНО-БЕЖЕВЫЙ', 'ТЕМНО-ЗЕЛЕНЫЙ', 'ТЕМНО-КОРИЧНЕВЫЙ', 'ТЕМНО-РОЗОВЫЙ', 'ТЕМНО-СЕРЫЙ', 'ТЕМНО-СИНИЙ', 'ТЕМНО-ФИОЛЕТОВЫЙ', 'ТЕРРАКОТОВЫЙ', 'ФИОЛЕТОВЫЙ',
    'ФУКСИЯ', 'ХАКИ', 'ЧЕРНЫЙ', 'ШОКОЛАДНЫЙ'
    ]
    
    const filePath = './public/moderation_marks/moderation_marks.html'

    // const fileContent = fs.readFileSync(filePath, 'utf-8')

    // const content = cio.load(fileContent)

    async function createImport(new_products) {
        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(i = 0; i < new_products.length; i++) {
            let size = ''            
                ws.getCell(`A${cellNumber}`).value = '6302'
                ws.getCell(`B${cellNumber}`).value = new_products[i]
                ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
                ws.getCell(`D${cellNumber}`).value = 'Артикул'
                ws.getCell(`E${cellNumber}`).value = vendorCodes[new_orders.indexOf(new_products[i])]
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

    async function getOrdersList() {

        let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {

            method: 'POST',
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':'144225',
                'Api-Key':'5d5a7191-2143-4a65-ba3a-b184958af6e8',
                'Content-Type':'application/json'
            },
            body: JSON.stringify({
                'dir': 'asc',
                'filter': {
                    'since':'2023-10-01T01:00:00.000Z',
                    'status':'awaiting_packaging',
                    'to':'2023-10-31T23:59:59.000Z'
                },
                'limit': 1000,
                'offset':0
            })

        })
        
        let result = await response.json()

        result.result.postings.forEach(e => {
            e.products.forEach(el => {
                if(new_orders.indexOf(el.name) < 0) {
                    if(el.name.indexOf('Набор махровых полотенец') >= 0 || el.name.indexOf('Гобелен') >= 0 || el.name.indexOf('Полотенце') >= 0 || el.name.indexOf('полотенце') >= 0 || el.name.indexOf('Постельное') >= 0 || el.name.indexOf('постельное') >= 0 || el.name.indexOf('Простыня') >= 0 || el.name.indexOf('Пододеяльник') >= 0 || el.name.indexOf('Наволочка') >= 0 || el.name.indexOf('Наматрасник') >= 0) {
                        new_orders.push(el.name)
                        vendorCodes.push(el.offer_id)
                    }
                } else {
                    console.log(el.name)
                }
            })
        })
        
    }
        // if(count === 1) {
        //     const divs = content('.details-cell_propsSecond_f-KWL')            
        //     divs.each((i, elem) => {
        //         // console.log(content(elem).text())
        //         let str = (content(elem).text()).trim()
        //         if(str.indexOf('Гобелен') >= 0 || str.indexOf('Полотенце') >= 0 || str.indexOf('полотенце') >= 0 || str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
        //     })
        // } else {
        //     for(i; i <= count; i++) {
        //         const divs = content('.details-cell_propsSecond_f-KWL')
        //         divs.each((i, elem) => {
        //             // console.log(content(elem).text())
        //             let str = (content(elem).text()).trim()
        //             if(str.indexOf('Постельное') >= 0 || str.indexOf('постельное') >= 0 || str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)
        //         })  
        //     }
        // }

    await getOrdersList()

    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'    

    wb.xlsx.readFile(fileName).then(() => {
        
        const ws = wb.getWorksheet('Краткий отчет')

        const c2 = ws.getColumn(2)

        c2.eachCell(c => {
           nat_cat.push(c.value)
        })

        const products = []
        const moderation_products = []

        const fileContent = fs.readFileSync(filePath, 'utf-8')

        const content = cio.load(fileContent)

        const spans = content('span')

        const divs = content('.gcJZKv')

        spans.each((i, elem) => {
            if(((content(elem).text()).indexOf('Гобеленовая') >= 0 || (content(elem).text()).indexOf('Полотенце') >= 0 || (content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
                products.push(content(elem).text())
            }
        })

        for(let i = 0; i < products.length; i++) {
            if(i%2 !== 0) {
                moderation_products.push(products[i])
            }
        }

        for(i = 0; i < new_orders.length; i++) {
            if(moderation_products.indexOf(new_orders[i].trim()) < 0 && nat_cat.indexOf(new_orders[i].trim()) < 0 && new_items.indexOf(new_orders[i].trim()) < 0){

                new_items.push(new_orders[i].trim())

            }

            if(nat_cat.indexOf(new_orders[i].trim()) >=0 && current_items.indexOf(new_orders[i].trim()) < 0) {

                current_items.push(new_orders[i].trim())

            }

            if(moderation_products.indexOf(new_orders[i].trim()) >=0 && current_items.indexOf(new_orders[i].trim()) < 0) {

                moderation_items.push(new_orders[i].trim())

            }

        }

        // console.log(moderation_products)

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

        moderation_items.forEach(elem => {
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-moderation">Модерируемый товар</span>
                     </div>`
        })

        html += `       </section>
                        <section class="action-form">
                            <button id="current-order"><a href="http://localhost:3030/ozon_marks_order" target="_blank">Создать заказ маркировки для актуальных товаров</a></button>
                            <button id="new-order"><a href="http://localhost:3030/ozon_new_marks_order" target="_blank">Создать заказ маркировки для новых товаров</a></button>
                        </section>
                        <div class="body-wrapper"></div>                        
                        ${footerComponent}`

        // html = '<h1 class="success">Import successfully done</h1>'
        res.send(html)

        if(new_items.length > 0) createImport(new_items)        

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
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    async function getOrdersList() {

        let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {

            method: 'POST',
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':'144225',
                'Api-Key':'5d5a7191-2143-4a65-ba3a-b184958af6e8',
                'Content-Type':'application/json'
            },
            body: JSON.stringify({
                'dir': 'asc',
                'filter': {
                    'since':'2023-10-01T01:00:00.000Z',
                    'status':'awaiting_packaging',
                    'to':'2023-10-31T23:59:59.000Z'
                },
                'limit': 1000,
                'offset':0
            })

        })
        
        let result = await response.json()

        result.result.postings.forEach(e => {
            e.products.forEach(el => {
                    if(el.name.indexOf('Набор махровых полотенец') >= 0 || el.name.indexOf('Гобелен') >= 0 || el.name.indexOf('Полотенце') >= 0 || el.name.indexOf('полотенце') >= 0 || el.name.indexOf('Постельное') >= 0 || el.name.indexOf('постельное') >= 0 || el.name.indexOf('Простыня') >= 0 || el.name.indexOf('Пододеяльник') >= 0 || el.name.indexOf('Наволочка') >= 0 || el.name.indexOf('Наматрасник') >= 0) {
                        if(new_orders.indexOf(el.name) < 0) {
                            new_orders.push(el.name)
                            quantity.push(el.quantity)
                        } else {
                            quantity[new_orders.indexOf(el.name)] += el.quantity
                        }
                    }
            })
        })
        
    }

    async function getQuantity() {
        let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {

            method: 'POST',
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':'144225',
                'Api-Key':'5d5a7191-2143-4a65-ba3a-b184958af6e8',
                'Content-Type':'application/json'
            },
            body: JSON.stringify({
                'dir': 'asc',
                'filter': {
                    'since':'2023-07-01T00:00:00.000Z',
                    'status':'awaiting_packaging',
                    'to':'2023-07-17T23:59:59.000Z'
                },
                'limit': 1000,
                'offset':0
            })

        })
        
        let result = await response.json()

        result.result.postings.forEach(e => {
            e.products.forEach(el => {                
                if(new_orders.indexOf(el.name) >= 0) {
                    quantity.push(el.quantity)
                }
            })
        })
    }

    await getOrdersList()

    // console.log(new_orders.length)

    // await getQuantity()

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

    html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Наименование</div>
                            <div class="header-cell">Статус</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

    // new_orders.forEach(elem => {
    //     if(nat_cat.indexOf(elem) < 0) {
    //         let index = new_orders.indexOf(elem)
    //         html += `<div class="table-row">
    //                     <span id="name">${elem}</span>
    //                     <span id="status-new">Новый товар</span>
    //                     <span id="quantity">${quantity[index]}</span>
    //                  </div>`
    //     }
    // })

    for(let i = 0; i < current_items.length; i++) {
        html += `<div class="table-row">
                    <span id="name">${current_items[i]}</span>
                    <span id="status-current">Актуальный товар</span>
                    <span id="quantity">${current_quantity[i]}</span>
                 </div>`
    }

    html += `</section>
             <div class="body-wrapper"></div>                        
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
            if(List[i].length > 0) {
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

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

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
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    // const fileContentOzon = fs.readFileSync(filePathOzon, 'utf-8')

    // const contentOzon = cio.load(fileContentOzon)

    async function getOrdersList() {

        let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {

            method: 'POST',
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':'144225',
                'Api-Key':'5d5a7191-2143-4a65-ba3a-b184958af6e8',
                'Content-Type':'application/json'
            },
            body: JSON.stringify({
                'dir': 'asc',
                'filter': {
                    'since':'2023-10-01T01:00:00.000Z',
                    'status':'awaiting_packaging',
                    'to':'2023-10-31T23:59:59.000Z'
                },
                'limit': 1000,
                'offset':0
            })

        })
        
        let result = await response.json()

        result.result.postings.forEach(e => {
            e.products.forEach(el => {
                    if(el.name.indexOf('Набор махровых полотенец') >= 0 || el.name.indexOf('Гобелен') >= 0 || el.name.indexOf('Полотенце') >= 0 || el.name.indexOf('полотенце') >= 0 || el.name.indexOf('Постельное') >= 0 || el.name.indexOf('постельное') >= 0 || el.name.indexOf('Простыня') >= 0 || el.name.indexOf('Пододеяльник') >= 0 || el.name.indexOf('Наволочка') >= 0 || el.name.indexOf('Наматрасник') >= 0) {
                        if(new_orders.indexOf(el.name) < 0) {
                            new_orders.push(el.name)
                            quantity.push(el.quantity)
                        } else {
                            quantity[new_orders.indexOf(el.name)] += el.quantity
                        }
                    }
            })
        })
        
    }

    // async function getQuantity() {

    //     let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {

    //         method: 'POST',
    //         headers: {
    //             'Host':'api-seller.ozon.ru',
    //             'Client-Id':'144225',
    //             'Api-Key':'5d5a7191-2143-4a65-ba3a-b184958af6e8',
    //             'Content-Type':'application/json'
    //         },
    //         body: JSON.stringify({
    //             'dir': 'asc',
    //             'filter': {
    //                 'since':'2023-07-01T00:00:00.000Z',
    //                 'status':'awaiting_packaging',
    //                 'to':'2023-07-17T23:59:59.000Z'
    //             },
    //             'limit': 1000,
    //             'offset':0
    //         })

    //     })
        
    //     let result = await response.json()

    //     result.result.postings.forEach(e => {
    //         e.products.forEach(el => {
    //                 if(el.name.indexOf('Гобелен') >= 0 || el.name.indexOf('Полотенце') >= 0 || el.name.indexOf('полотенце') >= 0 || el.name.indexOf('Постельное') >= 0 || el.name.indexOf('постельное') >= 0 || el.name.indexOf('Простыня') >= 0 || el.name.indexOf('Пододеяльник') >= 0 || el.name.indexOf('Наволочка') >= 0 || el.name.indexOf('Наматрасник') >= 0) {
    //                     quantity.push(el.quantity)
    //                 }
    //         })
    //     })
        
    // }

    await getOrdersList()

    // console.log(new_orders.length)

    // await getQuantity()

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    const spans = content('span')

    const divs = content('.gcJZKv')

    spans.each((i, elem) => {
        if(((content(elem).text()).indexOf('Гобеленовая') >= 0 || (content(elem).text()).indexOf('Полотенце') >= 0 || (content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
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

    html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Наименование</div>
                            <div class="header-cell">Статус</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

    new_orders.forEach(elem => {
        if(nat_cat.indexOf(elem) < 0) {
            let index = new_orders.indexOf(elem)
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-moderation">Модерируемый товар</span>
                        <span id="quantity">${quantity[index]}</span>
                     </div>`
            new_items.push(elem)
            new_quantity.push(quantity[index])
        }
    })

    html += `</section>
            <div class="body-wrapper"></div>
        ${footerComponent}`

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
                    if(List[i].length > 0) {
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
                                // console.log(moderation_gtins[moderation_products.indexOf(List[i][j])])
                                content += `<product>
                                                <gtin>${moderation_gtins[moderation_products.indexOf(List[i][j].trim())]}</gtin>
                                                <quantity>${Quantity[i][j]}</quantity>
                                                <serialNumberType>OPERATOR</serialNumberType>
                                                <cisType>UNIT</cisType>
                                                <templateId>10</templateId>
                                            </product>`
                            }
                        
                    content += `    </products>
                                </lp>
                            </order>`

                
                }
                    
                const date_ob = new Date()
                    
                let month = date_ob.getMonth() + 1
                    
                let filePath = ''
                    
                month < 10 ? filePath = `./public/orders/lp_ozon_new_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_ozon_new_${i}_${date_ob.getDate()}_${month}.xml`
                
                if(content !== '') {
                    fs.writeFileSync(filePath, content)
                }
                    
                content = ``
                
        }   
                
    }
                
    createOrder()

    res.send(html)

})

app.get('/wildberries', async function(req, res){
    
    const new_items = []
    const current_items = []
    const moderation_items = []
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

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    const filePath = './public/moderation_marks/moderation_marks.html'

    await wb.xlsx.readFile(hsFile)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    const products = []
    const moderation_products = []

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    const spans = content('span')

    const divs = content('.gcJZKv')

    spans.each((i, elem) => {
        if(((content(elem).text()).indexOf('Гобеленовая') >= 0 || (content(elem).text()).indexOf('Полотенце') >= 0 || (content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
            products.push(content(elem).text())
        }
    })

    for(let i = 0; i < products.length; i++) {
        if(i%2 !== 0) {
            moderation_products.push(products[i])
        }
    }

    await wb.xlsx.readFile(wbFile)

    const _ws = wb.getWorksheet('Сборочные задания')

    const c13 = _ws.getColumn(13)

    c13.eachCell(c => {
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
    console.log(moderation_items)

    testArray.forEach(elem => {
        if(moderation_products.indexOf(elem.trim()) < 0 && nat_cat.indexOf(elem.trim()) < 0 && new_items.indexOf(elem.trim()) < 0) {
            new_items.push(elem.trim())            
        }

        if(nat_cat.indexOf(elem.trim()) >= 0 && current_items.indexOf(elem.trim()) < 0) {
            current_items.push(elem.trim())
        }

        if(moderation_products.indexOf(elem.trim()) >= 0 && moderation_items.indexOf(elem.trim()) < 0) {
            moderation_items.push(elem.trim())
        }
    })

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    testArray.forEach(elem => {
        if(new_items.indexOf(elem) >= 0) {
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-new">Новый товар</span>
                     </div>`
        } else if(moderation_items.indexOf(elem) >= 0){
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-moderation">На модерации</span>
                     </div>`        
        } else {
            html += `<div class="table-row">
                        <span id="name">${elem}</span>
                        <span id="status-current">Актуальный товар</span>
                     </div>`
        }
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

    if(new_items.length > 0) {

        createImport(new_items)

    }

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
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    const c13 = _ws.getColumn(13)

    c13.eachCell(c => {
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
        if(orders.indexOf(el) < 0 && nat_cat.indexOf(el) >= 0) orders.push(el)
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

        if(_temp.length > 0) {
            orderList.push(_temp)
            _temp = []
        }

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < test_Array.length; i++) {
            
            if(nat_cat.indexOf(orders[i]) >= 0) {
                temp.push(test_Array[testArray.indexOf(orders[i])])
            }

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
            if(List[i].length > 0) {
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

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }
        
        // console.log(List)
        // console.log(Quantity)

        html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>
                        <div class="header-cell">Кол-во</div>
                    </div>
                <div class="header-wrapper"></div>`

        for(let i = 0; i < List.length; i++) {
            for(let j = 0; j < List[i].length; j++) {
                if(nat_cat.indexOf(List[i][j]) < 0) {
                    html += `<div class="table-row">
                                <span id="name">${List[i][j]}</span>
                                <span id="status-new">Новый товар</span>
                                <span id="quantity">${Quantity[i][j]}</span>
                             </div>`
                } else {
                    html += `<div class="table-row">
                                <span id="name">${List[i][j]}</span>
                                <span id="status-current">Актуальный товар</span>
                                <span id="quantity">${Quantity[i][j]}</span>
                             </div>`
                }
            }
        }

        html += `<section>
                <div class="body-wrapper"></div>
            ${footerComponent}`

        

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
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    const divs = content('.gcJZKv')

    spans.each((i, elem) => {
        if(((content(elem).text()).indexOf('Гобелен') >= 0 || (content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
            products.push(content(elem).text())
        }
    })

    // console.log(products)

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

    const c13 = _ws.getColumn(13)

    c13.eachCell(c => {
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
        if(orders.indexOf(el) < 0 && nat_cat.indexOf(el) < 0) orders.push(el)
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

            if(nat_cat.indexOf(orders[i]) < 0) {
                temp.push(test_Array[testArray.indexOf(orders[i])])
            }

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

        // console.log(List)

        // console.log(Quantity)

        html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>
                        <div class="header-cell">Кол-во</div>
                    </div>
                <div class="header-wrapper"></div>`

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
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
                        html += `<div class="table-row">
                                    <span id="name">${List[i][j]}</span>
                                    <span id="status-new">Новый товар</span>
                                    <span id="quantity">${Quantity[i][j]}</span>
                                </div>`
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

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_wb_new_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_wb_new_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }

        html += `<hr></div>
                    <section>${footerComponent}`

    }

    createOrder()

    res.send(html)

})

app.get('/sber', async function(req, res){    

    let objOrders = []
    // let names = []

    async function getShipments() {

        let orders = []

        let response = await fetch('https://api.megamarket.tech/api/market/v1/orderService/order/search', {
            method: 'POST',
            headers: {            
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                "meta": {},
                "data": {
                    "token": "A6B4E0AC-DD7F-4CF4-84D5-A772C59F38C4",
                    "dateFrom" : "2023-09-04T23:59:59Z",
                    "dateTo" : "2023-10-31T23:59:59Z",
                    "count": 100,
                    "statuses" : [
                        "CONFIRMED"
                    ]
                }            
            })
        })

        let result = await response.json()

        result.data.shipments.forEach(e => {
            orders.push(e)
        })

        return orders

    }    

    let orders = await getShipments()
    
    // console.log(result.data.shipments[0].items)

    // console.log(orders)

    async function getVendorCodes() {

        let vendorCodes = []
        let codesQuantity = []

        for(let i = 0; i < orders.length; i++) {
            
            let obj = {}
            let name = ''
            
            // console.log(typeof orders[i])

            let response = await fetch('https://api.megamarket.tech/api/market/v1/orderService/order/get', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'                    
                },
                body: JSON.stringify({
                    "data": {
                        "token": "A6B4E0AC-DD7F-4CF4-84D5-A772C59F38C4",
                        "shipments": [
                            `${orders[i].toString()}`
                        ]
                    },
                    "meta":{}
                })

            })

            let result = await response.json()

            // console.log(result)

            let products = []
            let productsQuantity = []

            result.data.shipments[0].items.forEach(e => {

                if(vendorCodes.indexOf(e.offerId) < 0) {

                    products.push(e.offerId)
                    vendorCodes.push(e.offerId)
                    codesQuantity.push(1)
                    productsQuantity.push(1)

                } else {

                    codesQuantity[vendorCodes.indexOf(e.offerId)] += 1
                    productsQuantity[products.indexOf(e.offerId)] += 1

                }

                obj = {

                    orderNumber: orders[i],
                    products: products,
                    productsQuantity: productsQuantity

                }

            })

            objOrders.push(obj)

        }

        // console.log(objOrders)
        return [vendorCodes, codesQuantity]

    }

    let [vendorCodes, codesQuantity] = await getVendorCodes()
    let markableVendorCodes = []

    async function getProductNames(array) {

        let productNames = []

        for(let i = 0; i < array.length; i++) {

            let response = await fetch('https://api-seller.ozon.ru/v2/product/info', {
                method: 'POST',
                headers: {
                    'Host': 'api-seller.ozon.ru',
                    'Client-Id': '144225',
                    'Api-Key': '5d5a7191-2143-4a65-ba3a-b184958af6e8',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    "offer_id": `${array[i]}`,
                    "product_id": 0,
                    "sku": 0
                })
            })

            let result = await response.json()

            if(result.result.name.indexOf('Подушка') < 0 && result.result.name.indexOf('Одеяло') < 0 && result.result.name.indexOf('Матрас') < 0) {
                if(productNames.indexOf(result.result.name) < 0) {
                    productNames.push(result.result.name)
                    markableVendorCodes.push(array[i])
                }
            }

        }

        return productNames

    }

    let products = await getProductNames(vendorCodes)

    // console.log(products)

    let html = ``

    async function beginRender() {

        html = `${headerComponent}
                        <title>Импорт - SBER</title>
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

    }

    await beginRender()

    const natCatFile = './public/Краткий отчет.xlsx'

    let nat_cat = []
    let gtins = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile(natCatFile)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c_1 = ws.getColumn(1)

    c_1.eachCell(c => {
        gtins.push(c.value)        
    })

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    let newProducts = []
    let moderProducts = []
    let currentProducts = []

    const filePath = './public/moderation_marks/moderation_marks.html'

    const _products = []
    const moderation_products = []
    const moderation_gtins = []

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    const spans = content('span')

    const divs = content('.gcJZKv')

    spans.each((i, elem) => {
        if(((content(elem).text()).indexOf('Гобеленовая') >= 0 || (content(elem).text()).indexOf('Полотенце') >= 0 || (content(elem).text()).indexOf('Постельное') >= 0 || (content(elem).text()).indexOf('Наволочка') >= 0 || (content(elem).text()).indexOf('Простыня') >= 0 || (content(elem).text()).indexOf('Пододеяльник') >= 0 || (content(elem).text()).indexOf('Наматрасник') >= 0 || (content(elem).text()).indexOf('Одеяло') >= 0 || (content(elem).text()).indexOf('Матрас') >= 0) && moderation_products.indexOf(content(elem).text()) < 0){
            _products.push(content(elem).text())
        }
    })

    for(let i = 0; i < _products.length; i++) {
        if(i%2 !== 0) {
            _products[i].replace('ё', 'е')
            moderation_products.push(_products[i])
        }
    }

    divs.each((i, elem) => {
        if((content(elem).text()).indexOf('029') >= 0) {
            moderation_gtins.push(content(elem).text())
        }
    })

    html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Номер заказа</div>
                            <div class="header-cell">Товары</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

    // console.log(nat_cat.indexOf('Пододеяльник из сатина 220 х 240 - М - 140 - Серый Люкс'))

    // console.log(2556 < 0)

    for(let i = 0; i < products.length; i++) {

        if(nat_cat.indexOf(products[i].trim()) < 0 && moderation_products.indexOf(products[i].trim()) < 0 && newProducts.indexOf(products[i].trim()) < 0) {
            newProducts.push(products[i])
        }
        
        if(nat_cat.indexOf(products[i].trim()) >= 0 && currentProducts.indexOf(products[i].trim()) < 0) {
            currentProducts.push(products[i])   
        }
        
        if(moderation_products.indexOf(products[i].trim()) >= 0 && moderProducts.indexOf(products[i].trim()) < 0 && currentProducts.indexOf(products[i].trim()) < 0) {
            moderProducts.push(products[i])
        }

    }

    // newProducts.forEach(e => {

    //     html += `<div class="table-row">
    //                 <span id="name">${e} - <span>${markableVendorCodes[products.indexOf(e)]}</span></span>
    //                 <span id="status-new">Новый товар</span>
    //             </div>`

    // })

    // currentProducts.forEach(e => {

    //     html += `<div class="table-row">
    //                 <span id="name">${e} - <span>${markableVendorCodes[products.indexOf(e)]}</span></span>
    //                 <span id="status-current">Актуальный товар</span>
    //              </div>`

    // })

    // moderProducts.forEach(e => {

    //     html += `<div class="table-row">
    //                 <span id="name">${e} - <span>${markableVendorCodes[products.indexOf(e)]}</span></span>
    //                 <span id="status-moderation">Модерируемый товар</span>
    //              </div>`

    // })

    const colors = ['БЕЖЕВЫЙ', 'БЕЛЫЙ', 'БИРЮЗОВЫЙ', 'БОРДОВЫЙ', 'БРОНЗОВЫЙ', 'ВАНИЛЬ', 'ВИШНЯ', 'ГОЛУБОЙ', 'ЖЁЛТЫЙ', 'ЗЕЛЁНЫЙ', 'ЗОЛОТОЙ', 'ИЗУМРУДНЫЙ',
    'КАПУЧИНО', 'КИРПИЧНЫЙ', 'КОРАЛЛОВЫЙ', 'КОРИЧНЕВЫЙ', 'КРАСНЫЙ', 'ЛАЙМ', 'ЛЕОПАРД', 'МАЛИНОВЫЙ', 'МЕДНЫЙ', 'МОЛОЧНЫЙ', 'МЯТНЫЙ', 'ОЛИВКОВЫЙ', 'ОРАНЖЕВЫЙ',
    'ПЕСОЧНЫЙ', 'ПЕРСИКОВЫЙ', 'ПУРПУРНЫЙ', 'РАЗНОЦВЕТНЫЙ', 'РОЗОВО-БЕЖЕВЫЙ', 'РОЗОВЫЙ', 'СВЕТЛО-БЕЖЕВЫЙ', 'СВЕТЛО-ЗЕЛЕНЫЙ', 'СВЕТЛО-КОРИЧНЕВЫЙ', 'СВЕТЛО-РОЗОВЫЙ',
    'СВЕТЛО-СЕРЫЙ', 'СВЕТЛО-СИНИЙ', 'СВЕТЛО-ФИОЛЕТОВЫЙ', 'СЕРЕБРЯНЫЙ', 'СЕРО-ЖЕЛТЫЙ', 'СЕРО-ГОЛУБОЙ', 'СЕРЫЙ', 'СИНИЙ', 'СИРЕНЕВЫЙ', 'ЛИЛОВЫЙ', 'СЛИВОВЫЙ',
    'ТЕМНО-БЕЖЕВЫЙ', 'ТЕМНО-ЗЕЛЕНЫЙ', 'ТЕМНО-КОРИЧНЕВЫЙ', 'ТЕМНО-РОЗОВЫЙ', 'ТЕМНО-СЕРЫЙ', 'ТЕМНО-СИНИЙ', 'ТЕМНО-ФИОЛЕТОВЫЙ', 'ТЕРРАКОТОВЫЙ', 'ФИОЛЕТОВЫЙ',
    'ФУКСИЯ', 'ХАКИ', 'ЧЕРНЫЙ', 'ШОКОЛАДНЫЙ'
    ]

    async function createImport(new_products) {

            const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
            const wb = new exl.Workbook()

            await wb.xlsx.readFile(fileName)

            const ws = wb.getWorksheet('IMPORT_TNVED_6302')

            let cellNumber = 5

            for(i = 0; i < new_products.length; i++) {
                let size = ''            
                    ws.getCell(`A${cellNumber}`).value = '6302'
                    ws.getCell(`B${cellNumber}`).value = new_products[i]
                    ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
                    ws.getCell(`D${cellNumber}`).value = 'Артикул'
                    ws.getCell(`E${cellNumber}`).value = markableVendorCodes[products.indexOf(new_products[i])]
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

            month < 10 ? filePath = `./public/sber/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_sber` : filePath = `./public/sber/IMPORT_TNVED_6302_${date_ob.getDate()}_${month}_sber`

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

    await createImport(newProducts)

    if(req.query.createOrder === undefined) {

        for(let i = 0; i < objOrders.length; i++) {
            
            html += `<div class="table-row">
            <span id="name">${objOrders[i].orderNumber}</span>`

            objOrders[i].products.forEach(e => {

                if(markableVendorCodes.indexOf(e) >= 0) {

                    html += `<span id="name-sber">${e} - ${products[markableVendorCodes.indexOf(e)]} ${objOrders[i].productsQuantity[objOrders[i].products.indexOf(e)]} шт.</span>`

                }            

            })

            html += `</div>`

        }

        html += `<div class="body-wrapper"></div>`

        html += `</section>
                <section class="action-form">
                    <button id="current-order"><a href="http://localhost:3030/sber?createOrder=current" target="_blank">Создать заказ маркировки для актуальных товаров</a></button>
                    <button id="new-order"><a href="http://localhost:3030/sber?createOrder=new" target="_blank">Создать заказ маркировки для новых товаров</a></button>
                </section>
                <div class="body-wrapper"></div>
            ${footerComponent}`

    }

    if(req.query.createOrder !== undefined) {

        // console.log(typeof req.query.createOrder)

        if(req.query.createOrder === 'new') {

            for(let i = 0; i < objOrders.length; i++) {
        
                html += `<div class="table-row">
                    <span id="name">${objOrders[i].orderNumber}</span>`

                    objOrders[i].products.forEach(e => {

                        if(markableVendorCodes.indexOf(e) >= 0) {

                            html += `<span id="name-sber">${e} - ${products[markableVendorCodes.indexOf(e)]} ${objOrders[i].productsQuantity[objOrders[i].products.indexOf(e)]} шт.</span>`

                        }            

                    })

                html += `</div>`

            }

            function createProductsList() {

                let orderList = []
                let _temp = []
        
                for (let i = 0; i < moderProducts.length; i++) {
        
                    _temp.push(moderProducts[i])
                    
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
        
                for(let i = 0; i < moderProducts.length; i++) {
        
                    temp.push(codesQuantity[vendorCodes.indexOf(markableVendorCodes[products.indexOf(moderProducts[i])])])
        
                        if(temp.length%10 === 0) {
                            quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                            temp = []
                        }
        
                }
        
                quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
        
                return quantityList

            }

            let productList = createProductsList()
            let quantityList = createQuantityList()

            // console.log([productList, quantityList])

            for(let i = 0; i < objOrders.length; i++) {
            
                html += `<div class="table-row">
                <span id="name">${objOrders[i].orderNumber}</span>`
    
                objOrders[i].products.forEach(e => {
    
                    if(markableVendorCodes.indexOf(e) >= 0) {
    
                        html += `<span id="name-sber">${e} - ${products[markableVendorCodes.indexOf(e)]} ${objOrders[i].productsQuantity[objOrders[i].products.indexOf(e)]} шт.</span>`
    
                    }            
    
                })
    
                html += `</div>`
    
            }
    
            html += `<div class="body-wrapper"></div>`

            html += `</section>
                <div class="body-wrapper"></div>
            ${footerComponent}`

            let content = ``

            for(let i = 0; i < productList.length; i++) {
                if(productList[i].length > 0) {
                    content += `<?xml version="1.0" encoding="utf-8"?>
                                        <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                            <lp>
                                                <productGroup>lp</productGroup>
                                                <contactPerson>333</contactPerson>
                                                <releaseMethodType>REMARK</releaseMethodType>
                                                <createMethodType>SELF_MADE</createMethodType>
                                                <productionOrderId>OZON</productionOrderId>
                                                <products>`

                    for(let j = 0; j < productList[i].length; j++) {
                        content += `<product>
                                        <gtin>${moderation_gtins[moderation_products.indexOf(productList[i][j].trim())]}</gtin>
                                        <quantity>${quantityList[i][j]}</quantity>
                                        <serialNumberType>OPERATOR</serialNumberType>
                                        <cisType>UNIT</cisType>
                                        <templateId>10</templateId>
                                    </product>`
                    }

                    content += `    </products>
                                </lp>
                            </order>`
                }

                const date_ob = new Date()
                    
                let month = date_ob.getMonth() + 1
                    
                let filePath = ''
                    
                month < 10 ? filePath = `./public/orders/lp_sber_new_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_sber_new_${i}_${date_ob.getDate()}_${month}.xml`
                
                if(content !== '') {
                    fs.writeFileSync(filePath, content)
                }
                    
                content = ``
            }

        }

        if(req.query.createOrder === 'current') {

            for(let i = 0; i < objOrders.length; i++) {
        
                html += `<div class="table-row">
                    <span id="name">${objOrders[i].orderNumber}</span>`

                    objOrders[i].products.forEach(e => {

                        if(markableVendorCodes.indexOf(e) >= 0) {

                            html += `<span id="name-sber">${e} - ${products[markableVendorCodes.indexOf(e)]} ${objOrders[i].productsQuantity[objOrders[i].products.indexOf(e)]} шт.</span>`

                        }            

                    })

                html += `</div>`

            }

            function createProductsList() {

                let orderList = []
                let _temp = []
        
                for (let i = 0; i < currentProducts.length; i++) {
        
                    _temp.push(currentProducts[i])
                    
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
        
                for(let i = 0; i < currentProducts.length; i++) {
        
                    temp.push(codesQuantity[vendorCodes.indexOf(markableVendorCodes[products.indexOf(currentProducts[i])])])
        
                        if(temp.length%10 === 0) {
                            quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                            temp = []
                        }
        
                }
        
                quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
        
                return quantityList

            }

            let productList = createProductsList()
            let quantityList = createQuantityList()

            // console.log([productList, quantityList])

            for(let i = 0; i < objOrders.length; i++) {
            
                html += `<div class="table-row">
                <span id="name">${objOrders[i].orderNumber}</span>`
    
                objOrders[i].products.forEach(e => {
    
                    if(markableVendorCodes.indexOf(e) >= 0) {
    
                        html += `<span id="name-sber">${e} - ${products[markableVendorCodes.indexOf(e)]} ${objOrders[i].productsQuantity[objOrders[i].products.indexOf(e)]} шт.</span>`
    
                    }            
    
                })
    
                html += `</div>`
    
            }
    
            html += `<div class="body-wrapper"></div>`

            html += `</section>
                <div class="body-wrapper"></div>
            ${footerComponent}`

            let content = ``

            for(let i = 0; i < productList.length; i++) {
                if(productList[i].length > 0) {
                    content += `<?xml version="1.0" encoding="utf-8"?>
                                        <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                            <lp>
                                                <productGroup>lp</productGroup>
                                                <contactPerson>333</contactPerson>
                                                <releaseMethodType>REMARK</releaseMethodType>
                                                <createMethodType>SELF_MADE</createMethodType>
                                                <productionOrderId>OZON</productionOrderId>
                                                <products>`

                    for(let j = 0; j < productList[i].length; j++) {
                        content += `<product>
                                        <gtin>0${gtins[nat_cat.indexOf(productList[i][j].trim())]}</gtin>
                                        <quantity>${quantityList[i][j]}</quantity>
                                        <serialNumberType>OPERATOR</serialNumberType>
                                        <cisType>UNIT</cisType>
                                        <templateId>10</templateId>
                                    </product>`
                    }

                    content += `    </products>
                                </lp>
                            </order>`
                }

                const date_ob = new Date()
                    
                let month = date_ob.getMonth() + 1
                    
                let filePath = ''
                    
                month < 10 ? filePath = `./public/orders/lp_sber_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_sber_${i}_${date_ob.getDate()}_${month}.xml`
                
                if(content !== '') {
                    fs.writeFileSync(filePath, content)
                }
                    
                content = ``
            }

        }

    }

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
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    html += `<div class="result">Файл remarking.xml успешно сформирован</div>
                <section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">КИЗ</div>
                            <div class="header-cell">Код ТНВЭД</div>
                            <div class="header-cell">Страна</div>
                        </div>
                        <div class="header-wrapper"></div>`

    marks.forEach(el => {
        if(el.length === 31) {
            html += `<div class="table-row">
                        <span type="text" id="mark">${el.replace(/</g, '&lt;')}</span>
                        <span id="name">6302100001</span>
                        <span id="name">РОССИЯ</span>
                     </div>`
        }
    })
    

    html += `   </div>
            </section>
        ${footerComponent}`

    res.send(html)
    
})

app.get('/sale_ozon', async function(req, res){

    const actualMarksFile = './public/actual_marks.xlsx'

    let html = `${headerComponent}
                    <title>Перемаркировка</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    const date_ob = new Date()

    let orders = []
    let consignments = []

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

    async function getActualList() {

        const [marks, status] = [[], []]

        await wb.xlsx.readFile(actualMarksFile)

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c16] = [ws.getColumn(1), ws.getColumn(16)]

        c1.eachCell(c => {
            marks.push(c.value)
        })

        c16.eachCell(c => {
            status.push(c.value)
        })

        const introduced_marks = []
        
        marks.forEach(e => {
            if(status[marks.indexOf(e)] == 'INTRODUCED') {
                introduced_marks.push(e)
            }
        })

    }

    await getActualList()

    //получаем данные из xlsx файла с реализациями и
    //формируем массив объектов реализаций

    async function getConsignments() {

        let consignments = []

        const consignmentDate = []

        const consignmentNumbers = []        

        const consignmentTypes = []

        const filePath = './public/distance/релизации.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('TDSheet')

        const [c2, c3, c8] = [ws.getColumn(2), ws.getColumn(3), ws.getColumn(8)]
        
        c2.eachCell(c => {
            let str = c.value
            consignmentDate.push(str.replace(str.substring(10), ''))
        })

        c3.eachCell(c => {
            let str = c.value
            consignmentNumbers.push(str.substring(str.length - 4))
        })

        c8.eachCell(c => {
            consignmentTypes.push(c.value)
        })

        let noRepeatConsignmentTypes = []

        for(let i = 0; i < consignmentTypes.length; i++) {
            if(consignmentTypes[i] != null && consignmentTypes[i].indexOf('ozon') >= 0 && noRepeatConsignmentTypes.indexOf(consignmentTypes[i]) < 0) {
                noRepeatConsignmentTypes.push(consignmentTypes[i])
            }
        }

        for(let i = 0; i < consignmentDate.length; i++) {
            let _tempArray = consignmentDate[i].split('.')
            let str = `${_tempArray[2]}-${_tempArray[1]}-${_tempArray[0]}`
            consignmentDate[i] = str
        }

        for(let i = 0; i < noRepeatConsignmentTypes.length; i++) {
            consignments.push({
                orderNumber: noRepeatConsignmentTypes[i].substring(5),
                consignmentNumber: consignmentNumbers[consignmentTypes.indexOf(noRepeatConsignmentTypes[i])],
                consignmentDate: consignmentDate[consignmentTypes.indexOf(noRepeatConsignmentTypes[i])]
            })
        }

        return consignments

    }

    consignments = await getConsignments()

    async function getOrders() {

        let orders = []

        let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {
            method: 'POST',
            headers: {
                'Host': 'api-seller.ozon.ru',
                'Client-Id': '144225',
                'Api-Key': '5d5a7191-2143-4a65-ba3a-b184958af6e8',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                'dir': 'asc',
                'filter':{
                    'since':'2023-07-14T08:00:00Z',
                    'to':'2023-08-26T08:00:00Z',
                    'status':'delivering'
                },
                'limit':1000,
                'offset':0
            })
        })
        
        let result = await response.json()
        
        result.result.postings.forEach(e => {
            let orderNumber = e.posting_number
            let products = []
            e.products.forEach(el => {
                let marks = []
                el.mandatory_mark.forEach(elem => {
                    marks.push(elem)
                })
                products.push({
                    name: el.name,
                    marksList: marks,
                    price: el.price
                })
            })

            let obj = {
                orderNumber: orderNumber,
                productsList: products
            }

            orders.push(obj)
            
        })

        return orders

    }

    orders = await getOrders()

    // orders.forEach(e => {
    //     console.log(e.orderNumber)
    // })

    let equals = []

    for(let i = 0; i < orders.length; i++) {
        for(let j = 0; j < consignments.length; j++) {
            if(orders[i].orderNumber == consignments[j].orderNumber) {
                equals.push(orders[i])
            }
        }
    }

    equals.forEach(e => {
        e.productsList.forEach(el => {
            if(el.marksList.length > 0) {
                if(el.marksList.indexOf('') < 0) {
                    for(let i = 0; i < el.marksList.length; i++) {
                        content += `<product>
                                        <cis><![CDATA[${el.marksList[i]}]]></cis>
                                        <cost>${(el.price).replace(el.price.substring(el.price.indexOf('.')), '')}00</cost>
                                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                                        <primary_document_number>${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentNumber}</primary_document_number>
                                        <primary_document_date>${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentDate}</primary_document_date>
                                    </product>`
                    }
                }
            }
        })
    })

    content += `</products_list>
            </withdrawal>`

    const fileName = `./public/distance/ozon_distance_${date_string}.xml`
    
    fs.writeFileSync(fileName, content)

    html += `<div class=result>Файл ${fileName.substring(fileName.lastIndexOf('/') + 1)} успешно сформирован</div>
            <section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">КИЗ</div>
                        <div class="header-cell">Цена</div>
                        <div class="header-cell">Тип документа</div>
                        <div class="header-cell">Номер документа</div>
                        <div class="header-cell">Дата документа</div>
                    </div>
                    <div class="header-wrapper"></div>`

    equals.forEach(e => {
        e.productsList.forEach(el => {
            if(el.marksList.length > 0) {
                if(el.marksList.indexOf('') < 0) {
                    for(let i = 0; i < el.marksList.length; i++) {
                        // console.log(el.marksList[i])
                        html += `<div class="table-row">
                                    <span type="text" id="mark">${el.marksList[i].replace(/</g, '&lt;')}</span>
                                    <span id="gtin">${(el.price).replace(el.price.substring(el.price.indexOf('.')), '')}00</span>
                                    <span id="name">CONSIGNMENT_NOTE</span>
                                    <span id="status">${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentNumber}</span>
                                    <span id="date">${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentDate}</span>
                                 </div>`
                    }
                }
            }
        })
    })

    html += `       </div>
                </section>
            ${footerComponent}`

    res.send(html)

})

app.get('/sale_wb', async function(req, res){

    let html = `${headerComponent}
                    <title>Перемаркировка</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`
    
    // let url = window.location.href
    // let str = url.split('/').reverse()[1]

    // document.title = str

    async function renderImportButtons(array) {

        let address = ''

        for(let i = 0; i < array.length; i++) {                
            if(array[i] === 'yandex') {
                address = 'yandex'
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                        </button>`
            }

            if(array[i] !== 'yandex') {
                array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                html += `<button class="button-import">
                            <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                        </button>`
            }
            
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

    const wbordersPath = './public/distance/wb_orders.xlsx'
    const consignmentsPath = './public/distance/релизации.xlsx'

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
    
    // let response = await fetch('https://suppliers-api.wildberries.ru/api/v3/orders?limit=10&next=0&dateFrom=1687755600&dateTo=1688187600',{
    //     method: 'GET',
    //     headers: {
    //         'Authorization':'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhY2Nlc3NJRCI6IjBhYmMxZWNmLTlmOWEtNDQzNi04YmNiLTM3Mjg1ZDJkYzJlZCJ9.-OGN5Jvwsf9XQHYy7LPPJjATV98xOSBXQMISSkjVNCg'
    //     }
    // })

    // let result = await response.json()

    // console.log(result.orders.forEach(e => {
    //     e.offices.forEach(el => {
    //         console.log(el)
    //     })
    // }))

    const wb = new exl.Workbook()

    let orders = []
    let consignments = []

    async function getOrders() {

        await wb.xlsx.readFile(wbordersPath)

        const ws = wb.getWorksheet('КИЗ')

        const orders = []

        const [orderNumbers, orderCises, orderPrices] = [[], [], []]

        const [c1, c3, c5] = [ws.getColumn(1), ws.getColumn(3), ws.getColumn(5)]
    
        c1.eachCell(c => {
            orderNumbers.push(c.value)
        })
    
        c3.eachCell(c => {
            orderCises.push(c.value)
        })
    
        c5.eachCell(c => {
            orderPrices.push(c.value)
        })

        for(let i = 0; i < orderNumbers.length; i++) {
            let obj = {
                orderNumber: orderNumbers[i],
                orderCis: orderCises[i],
                orderPrice: orderPrices[i]
            }

            orders.push(obj)
        }

        return orders

    }

    async function getConsignments() {

        await wb.xlsx.readFile(consignmentsPath)

        const ws = wb.getWorksheet('TDSheet')

        const [c2, c3, c8] = [ws.getColumn(2), ws.getColumn(3), ws.getColumn(8)]

        const [consDates, consNumbers, orderNumbers, wbNumbers] = [[], [], [], [], []]

        const numbers = []

        const consignments = []

        c2.eachCell(c => {
            let str = c.value.replace(c.value.substring(10), '')
            let date = str.split('.')
            consDates.push(`${date[2]}-${date[1]}-${date[0]}`)
        })

        c3.eachCell(c => {
            consNumbers.push(c.value.substring(c.value.length - 4))
        })

        c8.eachCell(c => {
            numbers.push(c.value)
            if(c.value != null) {
                wbNumbers.push(c.value)
                orderNumbers.push(c.value.substring(3))
            }
        })

        

        for(let i = 0; i < orderNumbers.length; i++) {

            let obj = {
                consDate: consDates[numbers.indexOf(wbNumbers[i])],
                consNumber: consNumbers[numbers.indexOf(wbNumbers[i])],
                orderNumber: orderNumbers[i]
            }

            consignments.push(obj)

        }

        // console.log(consignments)
        return consignments

    }

    orders = await getOrders()
    consignments = await getConsignments()

    // console.log(orders)
    // console.log(consignments)

    let equals = []

    for(let i = 0; i < orders.length; i++) {
        let index = consignments.indexOf(consignments.find(c => c.orderNumber == orders[i].orderNumber))
        if(index >= 0) {
            equals.push({
                consignmentNumber: consignments[index].consNumber,
                consignmentDate: consignments[index].consDate,
                consignmentPrice: orders[i].orderPrice,
                consignmentCis: orders[i].orderCis
            })
        }
    }

    for(let i = 0; i < equals.length; i++) {
        let price = ''

        if((equals[i].consignmentPrice.toString()).indexOf('.') >= 0) {
            let arr = (equals[i].consignmentPrice.toString()).split('.')
            price = arr[0]+arr[1]
        } else {
            price = equals[i].consignmentPrice + '00'
        }

        content += `<product>
                        <cis><![CDATA[${equals[i].consignmentCis}]]></cis>
                        <cost>${price}</cost>
                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                        <primary_document_number>${equals[i].consignmentNumber}</primary_document_number>
                        <primary_document_date>${equals[i].consignmentDate}</primary_document_date>
                    </product>`
        
    }
                            
    content += `</products_list>
            </withdrawal>`
    
    const fileName = `./public/distance/wb_distance_${date_string}.xml`
    
    fs.writeFileSync(fileName, content)

    html += `<div class=result>Файл ${fileName.substring(fileName.lastIndexOf('/') + 1)} успешно сформирован</div>
            <section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">КИЗ</div>
                        <div class="header-cell">Цена</div>
                        <div class="header-cell">Тип документа</div>
                        <div class="header-cell">Номер документа</div>
                        <div class="header-cell">Дата документа</div>
                    </div>
                    <div class="header-wrapper"></div>`

    for(let i = 0; i < equals.length; i++) {

        let price = ''

        if((equals[i].consignmentPrice.toString()).indexOf('.') >= 0) {
            let arr = (equals[i].consignmentPrice.toString()).split('.')
            price = arr[0]+arr[1]
        } else {
            price = equals[i].consignmentPrice + '00'
        }

        html += `<div class="table-row">
                    <span type="text" id="mark">${equals[i].consignmentCis.replace(/</g, '&lt;')}</span>
                    <span id="gtin">${price}</span>
                    <span id="name">CONSIGNMENT_NOTE</span>
                    <span id="status">${equals[i].consignmentNumber}</span>
                    <span id="date">${equals[i].consignmentDate}</span>
                </div>`

    }

    html += `           </div>
                    </section>
                <div class="body-wrapper"></div>
            ${footerComponent}`

    res.send(html)

})

app.get('/yandex', async function(req, res){

    if(req.query.cis === undefined) {

        let html = `${headerComponent}
                        <title>Я.Маркет</title>
                    </head>
                    <body>
                        ${navComponent}
                            <section class="sub-nav import-main">
                                <div class="import-control">`
        
        // let url = window.location.href
        // let str = url.split('/').reverse()[1]

        // document.title = str

        async function renderImportButtons(array) {

            let address = ''
    
            for(let i = 0; i < array.length; i++) {                
                if(array[i] === 'yandex') {
                    address = 'yandex'
                    html += `<button class="button-import">
                                <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                            </button>`
                }
    
                if(array[i] !== 'yandex') {
                    array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                    html += `<button class="button-import">
                                <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                            </button>`
                }
                
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

        html += `<div class="convert-form">
                    <h3 class="convert-form__header">Работа с КИЗ для Я.Маркета</h3>
                    <div class="body-wrapper"></div>
                    <div class="input-form">
                        <input class="input-form__input" type="text" placeholder="Введите КИЗ"/>
                        <a class="input-form__ref">Конвертировать</a>
                    </div>
                </div>`

        html += `${footerComponent}`

        res.send(html)

    }

    if(req.query.cis !== undefined) {

        let html = `${headerComponent}
                        <title>Перемаркировка</title>
                    </head>
                    <body>
                        ${navComponent}
                            <section class="sub-nav import-main">
                                <div class="import-control">`
        
        // let url = window.location.href
        // let str = url.split('/').reverse()[1]

        // document.title = str

        async function renderImportButtons(array) {

            let address = ''

            for(let i = 0; i < array.length; i++) {                
                if(array[i] === 'yandex') {
                    html += `<button class="button-import">
                                <a href="http://localhost:3030/${address}" target="_blank">Работа с ${array[i]}</a>
                            </button>`
                }

                if(array[i] !== 'yandex') {
                    array[i] === 'wb' ? address = 'wildberries' : address = array[i]
                    html += `<button class="button-import">
                                <a href="http://localhost:3030/${address}" target="_blank">Создать импорт для ${array[i]}</a>
                            </button>`
                }
                
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

        console.log(req.query.cis)

        let convertedString = req.query.cis.replace(req.query.cis.substring(0, 4), '')

        convertedString = convertedString.replace(/</g, '&lt;')

        convertedString = convertedString.replace(/&lt;GS>/g, '&bsol;u001d')

        convertedString = convertedString.replace(/ /g, '&plus;')

        html += `<div class="convert-form">
                    <h3 class="convert-form__header">Работа с КИЗ для Я.Маркета</h3>
                    <div class="body-wrapper"></div>
                    <div class="input-form">
                        <input class="input-form__input" type="text" placeholder="Введите КИЗ"/>
                        <a class="input-form__ref">Конвертировать</a>
                    </div>
                    <div class="body-wrapper"></div>
                    <div class="result-form">
                        <label class="result-form__label" for="result">Результат:</label>
                        <input name="result" class="result-form__input" type="text" value='${convertedString}'/>
                    </div>
                </div>`

        html += `<div class="body-wrapper"></div>${footerComponent}`

        res.send(html)

    }

})

app.listen(3030)