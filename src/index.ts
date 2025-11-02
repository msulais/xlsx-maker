import type { XLSXSheet } from "./sheets"

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
	password?: string
}

export class XLSX {
	#sheets = new Map<XLSXSheet['id'], XLSXSheet>()
	#options: XLSXOptions

	constructor(defaultSheet: XLSXSheet, options: XLSXOptions = {}) {
		this.#sheets.set(defaultSheet.id, defaultSheet)
		this.#options = options
	}

	get sheets() { return this
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