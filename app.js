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

    
    const filePath = './public/new_orders/new_orders.html'

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    const content = cio.load(fileContent)

    async function createImport(new_products) {
        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5
               
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

                //Вставка размера начало
                //Наволочки
                if(new_products[i].indexOf(' 40х40') >= 0 || new_products[i].indexOf(' 40 х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '40х40'
                else if(new_products[i].indexOf(' 40х60') >= 0 || new_products[i].indexOf(' 40 х 60') >= 0) ws.getCell(`K${cellNumber}`).value = '40х60'
                else if(new_products[i].indexOf(' 50х50') >= 0 || new_products[i].indexOf(' 50 х 50') >= 0) ws.getCell(`K${cellNumber}`).value = '50х50'
                else if(new_products[i].indexOf(' 60х60') >= 0 || new_products[i].indexOf(' 60 х 60') >= 0) ws.getCell(`K${cellNumber}`).value = '60х60'
                else if(new_products[i].indexOf(' 50х70') >= 0 || new_products[i].indexOf(' 50 х 70') >= 0) ws.getCell(`K${cellNumber}`).value = '50х70'
                else if(new_products[i].indexOf(' 70х70') >= 0 || new_products[i].indexOf(' 70 х 70') >= 0) ws.getCell(`K${cellNumber}`).value = '70х70'
                //Наматрасники
                else if(new_products[i].indexOf(' 60х120') >= 0 || new_products[i].indexOf(' 60 х 120') >= 0) ws.getCell(`K${cellNumber}`).value = '60х120'
                else if(new_products[i].indexOf(' 60х140') >= 0 || new_products[i].indexOf(' 60 х 140') >= 0) ws.getCell(`K${cellNumber}`).value = '60х140'
                else if(new_products[i].indexOf(' 70х120') >= 0 || new_products[i].indexOf(' 70 х 120') >= 0) ws.getCell(`K${cellNumber}`).value = '70х120'
                else if(new_products[i].indexOf(' 70х140') >= 0 || new_products[i].indexOf(' 70 х 140') >= 0) ws.getCell(`K${cellNumber}`).value = '70х140'
                else if(new_products[i].indexOf(' 70х200') >= 0 || new_products[i].indexOf(' 70 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200'
                else if(new_products[i].indexOf(' 80х200') >= 0 || new_products[i].indexOf(' 80 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200'
                else if(new_products[i].indexOf(' 90х200') >= 0 || new_products[i].indexOf(' 90 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200'
                else if(new_products[i].indexOf(' 120х200') >= 0 || new_products[i].indexOf(' 120 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200'
                else if(new_products[i].indexOf(' 140х200') >= 0 || new_products[i].indexOf(' 140 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200'
                else if(new_products[i].indexOf(' 150х220') >= 0 || new_products[i].indexOf(' 150 х 220') >= 0) ws.getCell(`K${cellNumber}`).value = '150х220'
                else if(new_products[i].indexOf(' 180х220') >= 0 || new_products[i].indexOf(' 180 х 220') >= 0) ws.getCell(`K${cellNumber}`).value = '180х220'
                else if(new_products[i].indexOf(' 160х200') >= 0 || new_products[i].indexOf(' 160 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200'
                else if(new_products[i].indexOf(' 170х200') >= 0 || new_products[i].indexOf(' 170 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200'
                else if(new_products[i].indexOf(' 180х200') >= 0 || new_products[i].indexOf(' 180 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200'
                else if(new_products[i].indexOf(' 200х200') >= 0 || new_products[i].indexOf(' 200 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200'
                else if(new_products[i].indexOf(' 200х220') >= 0 || new_products[i].indexOf(' 200 х 220') >= 0) ws.getCell(`K${cellNumber}`).value = '200х220'
                //Пододеяльники
                else if(new_products[i].indexOf(' 112х147') >= 0 || new_products[i].indexOf(' 112 х 147') >= 0) ws.getCell(`K${cellNumber}`).value = '112х147'
                else if(new_products[i].indexOf(' 145х215') >= 0 || new_products[i].indexOf(' 145 х 215') >= 0) ws.getCell(`K${cellNumber}`).value = '145х215'
                else if(new_products[i].indexOf(' 175х215') >= 0 || new_products[i].indexOf(' 175 х 215') >= 0) ws.getCell(`K${cellNumber}`).value = '175х215'
                else if(new_products[i].indexOf(' 200х200') >= 0 || new_products[i].indexOf(' 200 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200'
                else if(new_products[i].indexOf(' 220х240') >= 0 || new_products[i].indexOf(' 220 х 240') >= 0) ws.getCell(`K${cellNumber}`).value = '220х240'
                else if(new_products[i].indexOf(' 240х260') >= 0 || new_products[i].indexOf(' 240 х 260') >= 0) ws.getCell(`K${cellNumber}`).value = '240х260'
                else if(new_products[i].indexOf(' 150х200') >= 0 || new_products[i].indexOf(' 150 х 200') >= 0) ws.getCell(`K${cellNumber}`).value = '150х200'
                //Простыни
                else if(new_products[i].indexOf(' 70х200') >= 0 || new_products[i].indexOf(' 70 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '70х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '70х200'
                }
                else if(new_products[i].indexOf(' 80х200') >= 0 || new_products[i].indexOf(' 80 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '80х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '80х200'
                }
                else if(new_products[i].indexOf(' 90х200') >= 0 || new_products[i].indexOf(' 90 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '90х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '90х200'
                }
                else if(new_products[i].indexOf(' 120х200') >= 0 || new_products[i].indexOf(' 120 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '120х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '120х200'
                }
                else if(new_products[i].indexOf(' 140х200') >= 0 || new_products[i].indexOf(' 140 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '140х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '140х200'
                }
                else if(new_products[i].indexOf(' 160х200') >= 0 || new_products[i].indexOf(' 160 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '160х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '160х200'
                }
                else if(new_products[i].indexOf(' 170х200') >= 0 || new_products[i].indexOf(' 170 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '170х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '170х200'
                }
                else if(new_products[i].indexOf(' 180х200') >= 0 || new_products[i].indexOf(' 180 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '180х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '180х200'
                }
                else if(new_products[i].indexOf(' 200х200') >= 0 || new_products[i].indexOf(' 200 х 200') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '200х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '200х200'
                }
                else if(new_products[i].indexOf(' 210х210') >= 0 || new_products[i].indexOf(' 210 х 210') >= 0) {
                    if(new_products[i].indexOf('х10') >= 0) ws.getCell(`K${cellNumber}`).value = '210х210х10'
                    else if(new_products[i].indexOf('х20') >= 0) ws.getCell(`K${cellNumber}`).value = '210х210х20'
                    else if(new_products[i].indexOf('х30') >= 0) ws.getCell(`K${cellNumber}`).value = '210х210х30'
                    else if(new_products[i].indexOf('х40') >= 0) ws.getCell(`K${cellNumber}`).value = '210х210х40'
                    else if(new_products[i].indexOf('х 10') >= 0) ws.getCell(`K${cellNumber}`).value = '210х200х10'
                    else if(new_products[i].indexOf('х 20') >= 0) ws.getCell(`K${cellNumber}`).value = '210х200х20'
                    else if(new_products[i].indexOf('х 30') >= 0) ws.getCell(`K${cellNumber}`).value = '210х200х30'
                    else if(new_products[i].indexOf('х 40') >= 0) ws.getCell(`K${cellNumber}`).value = '210х200х40'
                    else ws.getCell(`K${cellNumber}`).value = '210х210'
                }
                //Вставка размера конец

                ws.getCell(`L${cellNumber}`).value = '6302100001'
                ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
                ws.getCell(`N${cellNumber}`).value = 'На модерации'                

                cellNumber++
            }
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

        month < 10 ? await wb.xlsx.writeFile(`./public/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}.xlsx`) : await wb.xlsx.writeFile(`./public/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}.xlsx`)

    }

    function getOrdersList(i, count) {
        if(count === 1) {
            const divs = content('.details-cell_propsSecond_f-KWL')
            divs.each((i, elem) => {
                // console.log(content(elem).text())
                let str = (content(elem).text()).trim()
                if(str.indexOf('Простыня') >= 0 || str.indexOf('Пододеяльник') >= 0 || str.indexOf('Наволочка') >= 0 || str.indexOf('Наматрасник') >= 0) new_orders.push(str)                
            })
        } else {
            for(i; i <= count; i++) {
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
            if(elem.indexOf('Постельное') < 0) html += `<p>${elem}</p>`
        })

        res.send(html)

        createImport(difference)

    }).catch(err => {
        console.log(err.message)
    })    

})

app.listen(3030)