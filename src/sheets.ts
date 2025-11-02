import { XLSXCell } from "./cells"

let SHEET_ID_COUNTER: number = 0

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

export class XLSXSheet {
	#name: string
	#id: number
	#order: number
	#cells: Map<XLSXCell['position'], XLSXCell>
	#page: XLSXPageSheet = {}
	#password?: string

	constructor(
		name: string,
		cells: XLSXCell[],
		options?: {
			order?: number,
			id?: number,
			page?: XLSXPageSheet
			password?: string
		}
	) {
		this.#name = name
		this.#order = options?.order ?? 0
		this.#id = options?.id ?? (++SHEET_ID_COUNTER)
		this.#cells = new Map(cells.map(v => [v.position, v]))
		this.#page = options?.page ?? {}
		this.#password = options?.password
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

	get password() {
		return this.#password
	}
	set password(value: string | undefined) {
		this.#password = value
	}

	get page() {
		return this.#page
	}
	set page(value: XLSXPageSheet) {
		this.#page = value
	}

	addCell(cell: XLSXCell) {
		this.#cells.set(cell.position, cell)
	}

	deleteCell(position: XLSXCell['position']) {
		this.#cells.delete(position)
	}

	getTable() {
		const rows: (XLSXCell | undefined)[][] = []
		const sortedCellsByColumn = (this
			.#cells.values().toArray()
			.sort((a, b) => a.coordinate[0] - b.coordinate[0])
		)
		for (const cell of sortedCellsByColumn) {
			const [x, y] = cell.coordinate
			let row = rows[y]
			if (!row) {
				row = rows[y] = []
			}

			row[x] = cell
 		}

		return rows
	}

	static copy(sheet: XLSXSheet) {
		return new XLSXSheet(sheet.#name, sheet.cells.map(v => XLSXCell.copy(v)), {
			id: sheet.#id,
			order: sheet.#order,
			page: structuredClone(sheet.#page),
			password: sheet.#password
		})
	}
}