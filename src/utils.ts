export function lettersToNumber(letters: string): number {
    let result = 0
    const base = 'A'.charCodeAt(0)
    for (let i = 0; i < letters.length; i++) {
        const currentChar = letters[i].toUpperCase()
        const charValue = currentChar.charCodeAt(0) - base + 1

        if (charValue < 1 || charValue > 26) {
            throw new Error(`Invalid character in input: ${letters[i]}`)
        }

        result = result * 26 + charValue
    }

    return result
}

export function dateDiffInDays(date1: Date, date2: Date): number {
	const MS_PER_DAY = 1000 * 60 * 60 * 24
	const utc1 = Date.UTC(date1.getFullYear(), date1.getMonth(), date1.getDate())
	const utc2 = Date.UTC(date2.getFullYear(), date2.getMonth(), date2.getDate())

	return Math.floor((utc2 - utc1) / MS_PER_DAY)
}