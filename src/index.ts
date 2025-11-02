import type { XLSXSheet } from "./sheets"
import { XMLElement } from "./xml"

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
		type FileName = string
		type XMLContent = string
		const files = new Map<FileName, XMLContent>()
		const xmlstart = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
		const options = this.#options
		const sheets = this.#sheets

		// TODO:
		function rels() {
			const xml = new XMLElement('Relationships',
				['xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships']
			)
			const xml_app = new XMLElement('Relationship',
				['Id', 'rId3'],
				['Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'],
				['Target', 'docProps/app.xml']
			)
			const xml_core = new XMLElement('Relationship',
				['Id', 'rId2'],
				['Type', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'],
				['Target', 'docProps/core.xml']
			)
			const xml_workbook = new XMLElement('Relationship',
				['Id', 'rId1'],
				['Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'],
				['Target', 'xl/workbook.xml']
			)
			xml.setChildren(xml_app, xml_core, xml_workbook)
			files.set('_rels/.rels', xml.toString())
		}

		function docProps() {
			function app() {
				const xml = new XMLElement('Properties',
					['xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'],
					['xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes']
				)
				const xml_application = new XMLElement('Application')
				const application = options.app?.name
				if (application) {
					xml_application.setChildren(application)
				}

				const xml_docSecurity = new XMLElement('DocSecurity')
				xml_docSecurity.setChildren(options.app?.docSecurity?.toString() ?? '0')

				const xml_scaleCrop = new XMLElement('ScaleCrop')
				xml_scaleCrop.setChildren(String(options.app?.scaleCrop ?? false))

				// TODO: <HeadingPairs>
				// TODO: <TitlesOfParts>

				const xml_company = new XMLElement('Company')
				const company = options.app?.company
				if (company) {
					xml_company.setChildren(company)
				}

				const xml_linksUpToDate = new XMLElement('LinksUpToDate')
				xml_linksUpToDate.setChildren(String(options.app?.linksUpToDate ?? false))

				const xml_sharedDoc = new XMLElement('SharedDoc')
				xml_sharedDoc.setChildren(String(options.app?.sharedDoc ?? false))

				const xml_hyperlinksChanged = new XMLElement('HyperlinksChanged')
				xml_hyperlinksChanged.setChildren(String(options.app?.hyperlinksChanged ?? false))

				const xml_appVersion = new XMLElement('AppVersion')
				xml_appVersion.setChildren(options.app?.version ?? '1')

				xml.setChildren(
					xml_appVersion, xml_docSecurity, xml_scaleCrop, xml_company, xml_linksUpToDate,
					xml_sharedDoc, xml_hyperlinksChanged, xml_appVersion
				)
				files.set('docProps/app.xml', xml.toString())
			}

			function core() {
				const xml = new XMLElement('cp:coreProperties',
					['xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'],
					['xmlns:dc', 'http://purl.org/dc/elements/1.1/'],
					['xmlns:dcterms', 'http://purl.org/dc/terms/'],
					['xmlns:dcmitype', 'http://purl.org/dc/dcmitype/'],
					['xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance'],
				)
				const xml_creator = new XMLElement('dc:creator')
				const creator = options.core?.creator
				if (creator) {
					xml_creator.setChildren(creator)
				}

				const xml_lastModifiedBy = new XMLElement('cp:lastModifiedBy')
				const lastModifiedBy = options.core?.lastModifiedBy
				if (lastModifiedBy || creator) {
					xml_lastModifiedBy.setChildren(lastModifiedBy ?? creator!)
				}

				const xml_created = new XMLElement('dcterms:created', ['xsi:type', 'dcterms:W3CDTF'])
				const created = options.core?.dateCreated
				if (created) {
					xml_created.setChildren(created.toISOString())
				}

				const xml_modified = new XMLElement('dcterms:modified', ['xsi:type', 'dcterms:W3CDTF'])
				const modified = options.core?.dateModified
				if (modified) {
					xml_modified.setChildren(modified.toISOString())
				}

				xml.setChildren(xml_creator, xml_lastModifiedBy, xml_created, xml_modified)
				files.set('docProps/core.xml', xml.toString())
			}

			app()
			core()
		}

		function xl() {
			function _rels() {
				const xml = new XMLElement('Relationships',
					['xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships']
				)

				// TODO: sheets
				const xml_styles = new XMLElement('Relationship',
					['Id', 'rId' + sheets.size + 1],
					['Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'],
					['Target', 'styles.xml']
				)
			}

			function worksheets() {
				// TODO
			}

			function styles() {}
			function sharedStrings() {}
			function workbook() {
				function _rels() {}
				function sheets() {}

				_rels()
				sheets()
			}

			_rels()
			worksheets()
			styles()
			sharedStrings()
			workbook()
		}

		function Content_Types() {}

		rels()
		docProps()
		xl()
		Content_Types()
	}
}