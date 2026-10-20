/**
 * @module util/zip
 * 极简 ZIP 读写（OOXML 容器）。只用 node:zlib，无第三方依赖。
 *
 * 为什么自己写：原 dsh-officecli 把整个文档读写外包给外部 officecli 可执行文件，
 * 一个 PATH/PE 解析问题就让全部 15 个工具瘫痪。OOXML 本身就是 ZIP + XML，
 * 进程内 100 行就能拿下，且能做到确定性输出（同样输入 -> 字节级相同）。
 */
import { deflateRawSync, inflateRawSync } from 'node:zlib';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const FLAG_UTF8 = 0x0800;

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

/** CRC32（ZIP 校验用）。 */
export function crc32(buf) {
  let c = 0 ^ -1;
  for (let i = 0; i < buf.length; i += 1) c = (c >>> 8) ^ CRC_TABLE[(c ^ buf[i]) & 0xff];
  return (c ^ -1) >>> 0;
}

/** 固定时间戳（1980-01-01 00:00），保证输出可复现。 */
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

/**
 * 打包为 ZIP Buffer。
 * @param {Array<{name:string, data:Buffer|string, store?:boolean}>} entries
 * @returns {Buffer}
 */
export function zipSync(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const entry of entries) {
    const raw = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(String(entry.data), 'utf8');
    const stored = entry.store === true;
    const body = stored ? raw : deflateRawSync(raw, { level: 9 });
    const crc = crc32(raw);
    const nameBuf = Buffer.from(entry.name, 'utf8');

    const local = Buffer.alloc(30);
    local.writeUInt32LE(SIG_LOCAL, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(FLAG_UTF8, 6);
    local.writeUInt16LE(stored ? 0 : 8, 8);
    local.writeUInt16LE(DOS_TIME, 10);
    local.writeUInt16LE(DOS_DATE, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, body);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(SIG_CENTRAL, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(FLAG_UTF8, 8);
    central.writeUInt16LE(stored ? 0 : 8, 10);
    central.writeUInt16LE(DOS_TIME, 12);
    central.writeUInt16LE(DOS_DATE, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(body.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);

    offset += local.length + nameBuf.length + body.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(SIG_EOCD, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, centralBuf, eocd]);
}

/**
 * 解包 ZIP，返回 name -> ()=>Buffer 的惰性读取器。
 * @param {Buffer} buf
 */
export function unzipSync(buf) {
  const eocdOffset = findEocd(buf);
  if (eocdOffset < 0) throw new Error('不是有效的 ZIP/OOXML 文件：找不到 End of Central Directory');
  const count = buf.readUInt16LE(eocdOffset + 10);
  let pointer = buf.readUInt32LE(eocdOffset + 16);
  const index = new Map();
  for (let i = 0; i < count; i += 1) {
    if (buf.readUInt32LE(pointer) !== SIG_CENTRAL) throw new Error('ZIP 中央目录损坏');
    const method = buf.readUInt16LE(pointer + 10);
    const compSize = buf.readUInt32LE(pointer + 20);
    const nameLen = buf.readUInt16LE(pointer + 28);
    const extraLen = buf.readUInt16LE(pointer + 30);
    const commentLen = buf.readUInt16LE(pointer + 32);
    const localOffset = buf.readUInt32LE(pointer + 42);
    const name = buf.toString('utf8', pointer + 46, pointer + 46 + nameLen);
    index.set(name, { method, compSize, localOffset });
    pointer += 46 + nameLen + extraLen + commentLen;
  }
  return {
    names: () => [...index.keys()],
    has: (name) => index.has(name),
    read(name) {
      const item = index.get(name);
      if (!item) throw new Error(`ZIP 中没有 ${name}`);
      const local = item.localOffset;
      if (buf.readUInt32LE(local) !== SIG_LOCAL) throw new Error('ZIP 本地头损坏');
      const nameLen = buf.readUInt16LE(local + 26);
      const extraLen = buf.readUInt16LE(local + 28);
      const start = local + 30 + nameLen + extraLen;
      const body = buf.subarray(start, start + item.compSize);
      if (item.method === 0) return Buffer.from(body);
      if (item.method === 8) return inflateRawSync(body);
      throw new Error(`不支持的压缩方法 ${item.method}`);
    },
    text(name) {
      return this.read(name).toString('utf8');
    },
  };
}

function findEocd(buf) {
  const maxBack = Math.min(buf.length, 0xffff + 22);
  for (let i = buf.length - 22; i >= buf.length - maxBack && i >= 0; i -= 1) {
    if (buf.readUInt32LE(i) === SIG_EOCD) return i;
  }
  return -1;
}
