import { dateDiffInDays, lettersToNumber } from "./utils"

let SHEET_ID_COUNTER: number = 0

export type XLSXCellValue = string | number | Date

export type XLSXCellBorder = {
	style?: string
	color?: `#${string}`
}

export type XLSXCellAttributes = {
	font?: {
		size?: number
		color?: `#${string}`
	}
	alignment?: {
		horizontal?: XLSXCellHorizontalAlign
		vertical?: XLSXCellVerticalAlign
		wrapText?: boolean
		shrinkToFit?: boolean
		textRotation?: number
		indent?: number
		readingOrder?: XLSXCellReadingOrder
	}
	fill?: `#${string}`
	border?: {
		left?: XLSXCellBorder
		right?: XLSXCellBorder
		top?: XLSXCellBorder
		bottom?: XLSXCellBorder
		diagonal?: XLSXCellBorder
	}
	format?: XLSXCellFormat
}

export enum XLSXCellVerticalAlign {
	top = 'top',
	center = 'center',
	bottom = 'bottom',
	justify = 'justify',
	distributed = 'distributed'
}

export enum XLSXCellHorizontalAlign {
	left = 'left',
	center = 'center',
	right = 'right',
	fill = 'fill',
	justify = 'justify',
	centerContinuous = 'centerContinuous',
	distributed = 'distributed'
}

export enum XLSXCellReadingOrder {
	default = 0,
	ltr = 1,
	rtl = 2
}

export enum XLSXCellFormat {
	// TODO:
}

export class XLSXCell {
	#position: string = 'A1'
	#value: XLSXCellValue = ''
	#absValue: string = ''
	#attributes: XLSXCellAttributes

	// coordinate start from (1,1) since that how what excel does (A1)
	#x: number = 1
	#y: number = 1

	constructor (
		position: string,
		value: XLSXCellValue,
		attributes?: XLSXCellAttributes
	){
		this.position = position
		this.value = value
		this.#attributes = attributes ?? {}
	}

	get absoluteValue() {
		return this.#absValue
	}

	get coordinate(): [x: number, y: number] {
		return [this.#x, this.#y]
	}

	get position() {
		return this.#position
	}

	set position (value: string) {
		value = value.replace(/\$/gs, '') // remove absolute position

		this.#position = /[A-Z]+?[1-9]+?[0-9]*/.test(value)? value : 'A1'
		let [x, y] = [1, 1]
		const X = this.#position.match(/^[A-Z]+/)
		const Y = this.#position.match(/[0-9]+$/)
		if (X) {
			x = Math.max(lettersToNumber(X[0]), 1)
		}
		if (Y) {
			y = Math.max(Number.parseInt(Y[0]), 1)
		}

		this.#x = x
		this.#y = y
	}

	get value() {
		return this.#value
	}

	set value(value: XLSXCellValue) {
		this.#value = value
		this.#updateAbsoluteValue()
	}

	#updateAbsoluteValue() {
		// TODO: consider cell format

		const v = this.#value
		if (v instanceof Date) {

			// NOTE: excel use "1900-01-01" as "1"
			const excelDateValue = dateDiffInDays(v, new Date(1900, 0, 1))
			this.#absValue = excelDateValue.toString()
			return
		}

		switch (typeof v) {
		case "number": return this.#absValue = (v as number).toString()
		case "string": return this.#absValue = v as string
		default: this.#absValue = String(v)
		}
	}

	static copy(value: XLSXCell) {
		return new XLSXCell(
			value.position,
			value.value,
			structuredClone(value.#attributes)
		)
	}
}

export class XLSXSheet {
	#name: string
	#id: number
	#order: number
	#cells: Map<XLSXCell['position'], XLSXCell>

	constructor(
		name: string,
		cells: XLSXCell[],
		options?: { order?: number, id?: number }
	) {
		this.#name = name
		this.#order = options?.order ?? 0
		this.#id = options?.id ?? (++SHEET_ID_COUNTER)
		this.#cells = new Map(cells.map(v => [v.position, v]))
	}

	get id() {
		return this.#id
	}

	get name() {
		return this.#name
	}
	set name(value: string) {
		this.#name = value
	}

	get order() {
		return this.#order
	}
	set order(value: number) {
		this.#order = value
	}

	get cells() {
		return this.#cells.values().toArray()
	}

	static copy(sheet: XLSXSheet) {
		return new XLSXSheet(sheet.#name, sheet.cells.map(v => XLSXCell.copy(v)), {
			id: sheet.#id,
			order: sheet.#order
		})
	}
}

export class XLSX {
	#sheets = new Map<XLSXSheet['id'], XLSXSheet>()

	constructor(defaultSheet: XLSXSheet) {
		this.#sheets.set(defaultSheet.id, defaultSheet)
	}

	get sheets() { return this
		.#sheets
		.values()
		.toArray()
		.sort((a, b) => a.order - b.order)
		.map(v => XLSXSheet.copy(v))
	}

	addSheet(sheet: XLSXSheet) {
		this.#sheets.set(sheet.id, sheet)
	}

	deleteSheet(sheetId: XLSXSheet['id']) {
		this.#sheets.delete(sheetId)
	}

	importFile() {
		// TODO:
	}

	exportFile() {
		// TODO:
	}
}