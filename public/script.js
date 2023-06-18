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