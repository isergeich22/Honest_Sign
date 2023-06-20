const express = require('express')
const tm = require('./example_app/components.js')
const data = require('./example_app/data.js')

const app = express()

app.get('/', async function(req, res) {
	const items = data.getItems()
	const html = tm.baseHtml(
		tm.navBar + tm.dynamic(tm.paginatedTable(items, 1)) + tm.footer
	)
	res.send(html)
})

app.get('/table', async function(req, res) {
	const items = data.getItems()
	res.send(tm.paginatedTable(items, req.query.page))
})

app.use(express.static(__dirname + '/public'))

app.listen(3030)