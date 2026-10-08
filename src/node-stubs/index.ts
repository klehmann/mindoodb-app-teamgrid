// Browser stand-ins for the Node built-ins the vendored xlsx-gateway imports
// at module level. Only its file helpers (atomic writes, SHA-256 inventory)
// use them, and the browser save path never calls those; a call here means a
// Node-only code path was reached and fails loudly.
function nodeOnly(name: string) {
  return (..._args: unknown[]): never => {
    throw new Error(`${name} is not available in the browser`)
  }
}

export const createHash = nodeOnly('crypto.createHash')
export const randomUUID = () => crypto.randomUUID()
export const closeSync = nodeOnly('fs.closeSync')
export const fsyncSync = nodeOnly('fs.fsyncSync')
export const openSync = nodeOnly('fs.openSync')
export const readFileSync = nodeOnly('fs.readFileSync')
export const writeFileSync = nodeOnly('fs.writeFileSync')
export const rename = nodeOnly('fs.rename')
export const rm = nodeOnly('fs.rm')
export const mkdtemp = nodeOnly('fs.mkdtemp')
export const stat = nodeOnly('fs.stat')
export const tmpdir = nodeOnly('os.tmpdir')
export const dirname = (path: string) => path.replace(/\/[^/]*$/, '') || '/'
export const join = (...parts: string[]) => parts.join('/').replace(/\/+/g, '/')
