
const home_button = document.querySelector('#home')
const import_button = document.querySelector('#import')
const cis_actions_button = document.querySelector('#cis_actions')
const import_control = document.querySelector('.import-control')
const marking_control = document.querySelector('.marking-control')
const nav_items = document.querySelectorAll('.nav-item')

home_button.addEventListener('click', () => {
    [import_control.style.display, marking_control.style.display] = ['none', 'none']
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    home_button.classList.add('active')
})

import_button.addEventListener('click', () => {
    import_control.style.display = 'flex'
    marking_control.style.display = 'none'
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    import_button.classList.add('active')
})

cis_actions_button.addEventListener('click', () => {
    marking_control.style.display = 'flex'
    import_control.style.display = 'none'
    nav_items.forEach(el => {
        el.classList.remove('active')
    })
    cis_actions_button.classList.add('active')
})

const table_header = document.querySelector('.marks-table-header')
const buttonTop = document.querySelector('#top')
const multipleList = document.querySelector('.multiple-list')
const statusList = document.querySelector('.status-list')
const statusRows = document.querySelectorAll('#status')

window.addEventListener('scroll', () => {

    if(scrollY > 100) {
        table_header.classList.add('--pinned')
        buttonTop.style.display = 'block'
    } else {
        table_header.classList.remove('--pinned')
        buttonTop.style.display = 'none'
    }

    buttonTop.addEventListener('click', () => {
        document.documentElement.scrollTop = 0
    })

})

multipleList.addEventListener('click', () => {

    const css = window.getComputedStyle(statusList)
    if(statusList.style.display == 'none' ||  css.display == 'none') {
        statusList.style.display = 'block'
    } else {
        statusList.style.display = 'none'
    }
    

})

statusRows.forEach(el => {
    if(el.innerHTML == 'В обороте') {
        el.style.color = '#36AD60'
    } else if(el.innerHTML == 'Нанесен') {
        el.style.color = 'rgb(240, 141, 27)'
    } else if(el.innerHTML == 'Выбыл') {
        el.style.color = 'rgb(122, 129, 155)'
    }
})