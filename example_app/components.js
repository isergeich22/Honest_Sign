exports.baseHtml = baseHtml = (contents) => `<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link rel="stylesheet" href="/css/styles.css" type="text/css">
        <link rel="shortcut icon" type="image/png" href="/favicon.png">
    </head>
    <body>
        ${contents}
        <script src="/script.js"></script>
    </body>
    </html>`

exports.navBar = navBar = `
    <header class="header">
        <nav>
            <img src="/img/chestnyj_znak.png" alt="честный знак">
            <p class="nav-item" id="home">Главная</p>
            <p class="nav-item" id="import">Создание импорт-файлов</p>
            <p class="nav-item" id="cis_actions">Действия с КИЗ</p>
        </nav>
    </header>`

exports.footer = footer =  `
    <button id="top" class="button-top">
    <svg width="24" height="24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <g clip-path="url(#ArrowLongUp_large_svg__clip0_35331_5070)">
            <path d="M12 2v20m0-20l7 6.364M12 2L5 8.364" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>
        </g><defs><clipPath id="ArrowLongUp_large_svg__clip0_35331_5070"><path fill="#fff" transform="rotate(90 12 12)" d="M0 0h24v24H0z">
        </path></clipPath></defs></svg>
    </button>`

exports.dynamic = dynamic = (contents) => `<div id="dynamic">${contents}</div>`

exports.table = table = (contents) => `
    <section class="table">
        <div class="marks-table">
            <div class="marks-table-header">
                <div class="header-cell">Наименование</div>
                <div class="header-cell">Статус</div>
            </div>
            <div class="header-wrapper"></div>
            ${contents}
        </div>
    </section>`

exports.tableRow = tableRow = (item) => `
    <div class="table-row">
        <span id="name">${item.name}</span>
        ${
            item.isNew
            ? '<span id="status-new">Новый товар</span>'
            : '<span id="status-current">Актуальный товар</span>'
        }
    </div>`

exports.pagination = pagination = (page, len) => `
   <div class="center">
        <div class="pagination">
            <a href="#" onclick="pagination(event, ${Math.max(1, page - 1)})">&laquo;</a>
            ${
                Array.from(Array(len+1).keys()).slice(1).map(a => `<a href="#" 
                        onclick="pagination(event, ${a})" 
                        ${a == page ? 'class="active"' : ''}>${a}
                    </a>`
                ).join("")
            }
            <a href="#" onclick="pagination(event, ${Math.min(len, page + 1)})">&raquo;</a>
        </div>
    </div>
`

exports.paginatedTable = paginatedTable = (items, page=1) => {
    page = Number(page)
    const pageSize = 10
    const start = (page - 1) * pageSize
    const end = start + pageSize
    slice = items.slice(start, Math.min(end, items.length))
    return table(
        slice.map(tableRow).join("") + pagination(page, Math.ceil(items.length/pageSize))
    )
}