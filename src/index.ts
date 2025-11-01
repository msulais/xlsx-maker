import { dateDiffInDays, lettersToNumber } from "./utils"

let SHEET_ID_COUNTER: number = 0

export type AARRGGBB = number
export type XLSXCellValue = string | number | Date

export type XLSXCellAttributes = {
	borderBottomColor?: AARRGGBB
	borderBottomStyle?: XLSXCellBorderStyle
	borderLeftColor?: AARRGGBB
	borderLeftStyle?: XLSXCellBorderStyle
	borderRightColor?: AARRGGBB
	borderRightStyle?: XLSXCellBorderStyle
	borderTopColor?: AARRGGBB
	borderTopStyle?: XLSXCellBorderStyle
	color?: AARRGGBB
	fill?: AARRGGBB
	fontSize?: number
	format?: XLSXCellFormat | string
	hidden?: boolean
	horizontalAlign?: XLSXCellHorizontalAlign
	indent?: number
	locked?: boolean
	readingOrder?: XLSXCellReadingOrder
	shrinkToFit?: boolean
	textRotation?: number
	verticalAlign?: XLSXCellVerticalAlign
	wrapText?: boolean
}

export enum XLSXCellBorderStyle {
	dashed = 'dashed',
	dotted = 'dotted',
	double = 'double',
	medium = 'medium',
	thick = 'thick',
	thin = 'thin',
}

export enum XLSXCellVerticalAlign {
	bottom = 'bottom',
	distributed = 'distributed',
	center = 'center',
	justify = 'justify',
	top = 'top',
}

export enum XLSXCellHorizontalAlign {
	left = 'left',
	center = 'center',
	right = 'right',
	fill = 'fill',
	justify = 'justify',
	centerContinuous = 'centerContinuous',
	distributed = 'distributed',
}

export enum XLSXCellReadingOrder {
	default = 0,
	ltr = 1,
	rtl = 2
}

export enum XLSXCellFormat {
	/** ## Format = `General` */
	general = 0,

	/** ## Format = `0` */
	number1 = 1,

	/** ## Format = `0.00` */
	number2 = 2,

	/** ## Format = `#.##0` */
	number3 = 3,

	/** ## Format = `#,##0.00` */
	number4 = 4,

	/** ## Format = `$#,##0;($#,##0)` */
	currency1 = 5,

	/** ## Format = `$#,##0;[Red]($#,##0)` */
	currency2 = 6,

	/** ## Format = `$#,##0.00;($#,##0.00)` */
	currency3 = 7,

	/** ## Format = `$#,##0.00;[Red]($#,##0.00)` */
	currency4 = 8,

	/** ## Format = `0%` */
	percentage1 = 9,

	/** ## Format = `0.00%` */
	percentage2 = 10,

	/** ## Format = `0.00E+00` */
	scientific = 11,

	/** ## Format = `# ?/?` */
	fraction1 = 12,

	/** ## Format = `# ??/??` */
	fraction2 = 13,

	/** ## Format = `d/m/yyyy` */
	date1 = 14,

	/** ## Format = `d-mmm-yy` */
	date2 = 15,

	/** ## Format = `d-mmm` */
	date3 = 16,

	/** ## Format = `mmm-yy` */
	date4 = 17,

	/** ## Format = `h:mm tt` */
	time1 = 18,

	/** ## Format = `h:mm:ss tt` */
	time2 = 19,

	/** ## Format = `H:mm` */
	time3 = 20,

	/** ## Format = `H:mm:ss` */
	time4 = 21,

	/** ## Format = `mm:ss` */
	time5 = 45,

	/** ## Format = `[h]:mm:ss` */
	time6 = 46,

	/** ## Format = `mmss.0` */
	time7 = 47,

	/** ## Format = `m/d/yy h:mm` */
	datetime = 22,

	/** ## Format = `#,##0 ;(#,##0)` */
	accounting1 = 37,

	/** ## Format = `#,##0 ;[Red](#,##0)` */
	accounting2 = 38,

	/** ## Format = `#,##0.00;(#,##0.00)` */
	accounting3 = 39,

	/** ## Format = `#,##0.00;[Red](#,##0.00)` */
	accounting4 = 40,

	/** ## Format = `@` */
	text = 49,
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