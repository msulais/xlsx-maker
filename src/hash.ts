interface XLSXHashData {
	algorithmName: string
	hashValue: string
	saltValue: string
	spinCount: string
}

const subtleCrypto: SubtleCrypto = window.crypto.subtle

function strToUtf16le(str: string): Uint8Array {
	const buffer = new ArrayBuffer(str.length * 2)
	const view = new DataView(buffer)
	for (let i = 0; i < str.length; i++) {
		view.setUint16(i * 2, str.charCodeAt(i), true)
	}

	return new Uint8Array(buffer)
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
	const byteArray = new Uint8Array(buffer)
	const byteString = String.fromCharCode.apply(null, Array.from(byteArray))
	return btoa(byteString)
}

function concatArrayBuffers(arr1: ArrayBuffer, arr2: ArrayBuffer): Uint8Array {
	const tmp = new Uint8Array(arr1.byteLength + arr2.byteLength)
	tmp.set(new Uint8Array(arr1), 0)
	tmp.set(new Uint8Array(arr2), arr1.byteLength)
	return tmp
}

export async function generateXLSXHashProtection(password: string): Promise<XLSXHashData> {
	const spinCount = 100000
	const algorithmName = 'SHA-512'
	const saltBytes = window.crypto.getRandomValues(new Uint8Array(16))
	const passwordBytes = strToUtf16le(password)
	const initialBuffer = concatArrayBuffers(
		saltBytes.buffer,
		passwordBytes.buffer as ArrayBuffer
	)
	let currentHash: ArrayBuffer = await subtleCrypto.digest(
		algorithmName,
		initialBuffer as Uint8Array<ArrayBuffer>
	)
	for (let i = 0; i < spinCount; i++) {
		const iterBuffer = new ArrayBuffer(4)
		new DataView(iterBuffer).setUint32(0, i, true)
		const combinedBuffer = concatArrayBuffers(currentHash, iterBuffer)
		currentHash = await subtleCrypto.digest(algorithmName, combinedBuffer as BufferSource)
	}

	const hashValue = arrayBufferToBase64(currentHash)
	const saltValue = arrayBufferToBase64(saltBytes.buffer)

	return {
		algorithmName,
		hashValue,
		saltValue,
		spinCount: spinCount.toString(),
	}
}