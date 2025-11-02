import { dateDiffInDays, lettersToNumber } from "./utils"

let SHEET_ID_COUNTER: number = 0

export type AARRGGBB = number
export type XLSXCellValue = string | number | Date

export type XLSXPageSheet = {
	printArea?: `${XLSXCell['position']}:${XLSXCell['position']}`
	pictureUrl?: string
	margins?: {
		bottom?: number
		footer?: number
		header?: number
		left?: number
		right?: number
		top?: number
	}
	setup?: {
		blackAndWhite?: boolean
		cellComments?: 'none' | 'asDisplayed' | 'atEnd'
		copies?: number
		draft?: boolean
		errors?: 'displayed' | 'blank' | 'dash' | 'na'
		firstPageNumber?: number
		fitToHeight?: number
		fitToWidth?: number
		horizontalDpi?: number
		orientation?: 'default' | 'portrait' | 'landscape'
		pageOrder?: 'downThenOver' | 'overThenDown'
		paperSize?: XLSXPageSize
		scale?: number
		useFirstPageNumber?: boolean
		usePrinterDefaults?: boolean
		verticalDpi?: number
	}
}

export type XLSXOptions = {
	app?: {
		company?: string
		docSecurity?: number
		hyperlinksChanged?: boolean
		linksUpToDate?: boolean
		name?: string
		scaleCrop?: boolean
		sharedDoc?: boolean
		version?: string
	}
	core?: {
		creator?: string
		dateCreated?: Date
		dateModified?: Date
		lastModifiedBy?: string
	}
}

export type XLSXCellAttributes = {
	borderBottomColor?: AARRGGBB
	borderBottomStyle?: 'dashed' | 'dotted' | 'dobule' | 'medium' | 'thick' | 'thin'
	borderLeftColor?: AARRGGBB
	borderLeftStyle?: 'dashed' | 'dotted' | 'dobule' | 'medium' | 'thick' | 'thin'
	borderRightColor?: AARRGGBB
	borderRightStyle?: 'dashed' | 'dotted' | 'dobule' | 'medium' | 'thick' | 'thin'
	borderTopColor?: AARRGGBB
	borderTopStyle?: 'dashed' | 'dotted' | 'dobule' | 'medium' | 'thick' | 'thin'
	color?: AARRGGBB
	fill?: AARRGGBB
	fontSize?: number
	format?: XLSXCellFormat | string
	hidden?: boolean
	horizontalAlign?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
	indent?: number
	locked?: boolean
	readingOrder?: XLSXCellReadingOrder
	shrinkToFit?: boolean
	textRotation?: number
	verticalAlign?: 'bottom' | 'distributed' | 'center' | 'justify' | 'top'
	wrapText?: boolean
}

export enum XLSXPageSize {
	A4 = 9,
	A6 = 70,
	A5 = 11,
	A3 = 8,
	A2 = 66,
	A3Plus = 67,
	size10x15 = 45, // Photo 10x15 cm (or 4x6 in)
	size13x18 = 46, // Photo 13x18 cm (or 5x7 in)
	size9x13 = 47, // Photo 9x13 cm (or 3.5x5 in)
	size5x8 = 48, // 5x8 in (Statement size is 6)
	size20x25 = 49, // 20x25 cm (or 8x10 in)
	letter = 1,
	legal = 5,
	size8x13 = 41, // Folio or German Legal Fanfold
	indianLegal = 70, // Often mapped to A6/A4 or similar custom sizes in non-MS systems
	envelope10 = 20, // Envelope #10 (4 1/8 x 9 1/2 in)
	envelopeDL = 27, // Envelope DL (110 x 220 mm)
	envelopeC6 = 31, // Envelope C6 (114 x 162 mm)
	B4 = 12,
	B3 = 68,
	size8K = 65, // A common Asian paper standard
	size16K = 64, // A common Asian paper standard
	size100x148 = 44, // 100x148 mm (A6 photo card size)
	wide16x9 = 71,
	custom = 256, // TODO: how to implement this?
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

	constructor(
		position: string,
		value: XLSXCellValue,
		attributes?: XLSXCellAttributes
	) {
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

	set position(value: string) {
		value = value.replace(/\$/gs, '') // remove absolute position

		this.#position = /[A-Z]+?[1-9]+?[0-9]*/.test(value) ? value : 'A1'
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
	#page: XLSXPageSheet = {}

	constructor(
		name: string,
		cells: XLSXCell[],
		options?: {
			order?: number,
			id?: number,
			page?: XLSXPageSheet
		}
	) {
		this.#name = name
		this.#order = options?.order ?? 0
		this.#id = options?.id ?? (++SHEET_ID_COUNTER)
		this.#cells = new Map(cells.map(v => [v.position, v]))
		this.#page = options?.page ?? {}
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

	addCell(cell: XLSXCell) {
		this.#cells.set(cell.position, cell)
	}

	deleteCell(position: XLSXCell['position']) {
		this.#cells.delete(position)
	}

	static copy(sheet: XLSXSheet) {
		return new XLSXSheet(sheet.#name, sheet.cells.map(v => XLSXCell.copy(v)), {
			id: sheet.#id,
			order: sheet.#order,
			page: structuredClone(sheet.#page)
		})
	}
}

export class XLSX {
	#sheets = new Map<XLSXSheet['id'], XLSXSheet>()
	#options: XLSXOptions

	constructor(defaultSheet: XLSXSheet, options: XLSXOptions = {}) {
		this.#sheets.set(defaultSheet.id, defaultSheet)
		this.#options = options
	}

	get sheets() {
		return this
			.#sheets
			.values()
			.toArray()
			.sort((a, b) => a.order - b.order)
	}

	get options() {
		return this.#options
	}

	set options(value: XLSXOptions) {
		this.#options = value
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