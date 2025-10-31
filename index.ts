type SheetId = number

let SHEET_ID_COUNTER: number = 0

export class XSLXFormula {
	constructor (){}
}

export class XLSXCell {
	private _sheetId: SheetId
	private _position: string
	private _value: string | XSLXFormula

	constructor (sheetId: SheetId, position: string, value: string | XSLXFormula){
		this._position = /[A-Z]+?[0-9]+/.test(position)? position : 'A1'
		this._value = value
		this._sheetId = sheetId
	}

	get isFormula() {
		return this._value instanceof XSLXFormula
	}

	get value() {
		return this._value
	}

	get position() {
		return this._position
	}

	get sheetId() {
		return this._sheetId
	}
}

export class XLSX {
	private _data = new Map<XLSXCell['position'], XLSXCell>()
	private _sheets = new Map<SheetId, {id: SheetId, name: string, order: number}>()

	constructor(sheetName: string) {
		const id = this.addSheet(sheetName)
	}

	get sheets() { return this
		._sheets
		.values()
		.toArray()
		.sort((a, b) => a.order - b.order)
		.map(v => ({...v}))
	}

	updateCell(cell: XLSXCell) {
		this._data.set(cell.position, cell)
	}

	deleteCell(cell: XLSXCell) {
		this._data.delete(cell.position)
	}

	getCell(position: XLSXCell['position']) {
		return this._data.get(position)
	}

	addSheet(sheetName: string, order: number = this._sheets.size + 1): number {
		const id = ++SHEET_ID_COUNTER
		this._sheets.set(id, {id, name: sheetName, order})
		return 0
	}

	deleteSheet(sheetId: SheetId) {
		this._sheets.delete(sheetId)
	}

	clearSheet(sheetId: SheetId) {
		for (const cell of this._data) {
			if (cell[1].sheetId !== sheetId) {
				continue
			}

			this._data.delete(cell[1].position)
		}
	}

	getTable() {
		// TODO
	}

	importFile() {
		// TODO:
	}

	exportFile() {
		// TODO:
	}
}