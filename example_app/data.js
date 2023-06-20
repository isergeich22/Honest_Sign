let items  = []

exports.getItems = getItems = () => {
	if (!items.length) {
		items = Array.from(Array(100).keys()).map(a => ({
			isNew: Math.random() < 0.1 ? true : false,
			name: `bimbom ${a}`
		})).sort((a, b) => b.isNew - a.isNew)
	}
	return items
}

exports.getItems = getItems