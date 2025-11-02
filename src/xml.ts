export class XMLElement {
	#tagName: string
	#attributes: Map<string, string>
	#children: (XMLElement | string)[] = []

	constructor (tagName: string, ...attributes: [name: string, value: string][]) {
		this.#tagName = tagName
		this.#attributes = new Map(attributes)
	}

	#sanitizeInput(input: string) {
		// TODO: sanitize
		return input
	}

	setChildren(...xmls: (XMLElement | string)[]) {
		this.#children = xmls
	}

	setAttribute(name: string, value: string) {
		this.#attributes.set(name, value)
	}

	toString() {
		let xml = '<' + this.#tagName
		for (const attribute of this.#attributes) {
			xml += ` ${attribute[0]}="${this.#sanitizeInput(attribute[1])}"`
		}

		if (this.#children.length <= 0) {
			return xml + '/>'
		}

		xml += '>'
		for (const child of this.#children) {
			xml += (child instanceof XMLElement
				? child.toString()
				: this.#sanitizeInput(child)
			)
		}

		return xml + '</' + this.#tagName + '>'
	}
}