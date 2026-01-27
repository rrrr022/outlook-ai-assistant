const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const sizes = [16, 32, 64, 80];
const assetsDir = path.join(__dirname, '..', 'src', 'client', 'assets');

// Ensure assets directory exists
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

// CRC32 implementation
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) {
    c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
  }
  crcTable[n] = c >>> 0;
}

function crc32(data) {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc = crcTable[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  }
  return ((crc ^ 0xffffffff) >>> 0);
}

function createChunk(type, data) {
  const typeBytes = Buffer.from(type, 'ascii');
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  
  const toChecksum = Buffer.concat([typeBytes, data]);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(toChecksum));
  
  return Buffer.concat([length, typeBytes, data, checksum]);
}

function createPNG(width, height, r, g, b) {
  // PNG signature
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  
  // IHDR data
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);
  ihdrData.writeUInt32BE(height, 4);
  ihdrData[8] = 8;  // bit depth
  ihdrData[9] = 2;  // color type RGB
  ihdrData[10] = 0; // compression
  ihdrData[11] = 0; // filter
  ihdrData[12] = 0; // interlace
  
  // Raw image data (filter byte + RGB for each row)
  const rawData = [];
  for (let y = 0; y < height; y++) {
    rawData.push(0); // filter byte (none)
    for (let x = 0; x < width; x++) {
      rawData.push(r, g, b);
    }
  }
  
  const compressed = zlib.deflateSync(Buffer.from(rawData), { level: 9 });
  
  // Build PNG
  const ihdr = createChunk('IHDR', ihdrData);
  const idat = createChunk('IDAT', compressed);
  const iend = createChunk('IEND', Buffer.alloc(0));
  
  return Buffer.concat([signature, ihdr, idat, iend]);
}

// Generate icons with blue color (Microsoft blue #0078D4)
sizes.forEach(size => {
  const png = createPNG(size, size, 0, 120, 212);
  const filename = `icon-${size}.png`;
  fs.writeFileSync(path.join(assetsDir, filename), png);
  console.log(`Created ${filename}`);
});

console.log('✅ All icons created in', assetsDir);
