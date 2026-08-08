const fs = require('fs');
const zlib = require('zlib');

// Font pixel patterns for "A I M S" and "H R I S"
// 5x7 bitmap font definitions
const fontMap = {
  'A': [
    [0,1,1,0],
    [1,0,0,1],
    [1,1,1,1],
    [1,0,0,1],
    [1,0,0,1]
  ],
  'I': [
    [1,1,1],
    [0,1,0],
    [0,1,0],
    [0,1,0],
    [1,1,1]
  ],
  'M': [
    [1,0,0,0,1],
    [1,1,0,1,1],
    [1,0,1,0,1],
    [1,0,0,0,1],
    [1,0,0,0,1]
  ],
  'S': [
    [0,1,1,1],
    [1,0,0,0],
    [0,1,1,0],
    [0,0,0,1],
    [1,1,1,0]
  ],
  'H': [
    [1,0,0,1],
    [1,0,0,1],
    [1,1,1,1],
    [1,0,0,1],
    [1,0,0,1]
  ],
  'R': [
    [1,1,1,0],
    [1,0,0,1],
    [1,1,1,0],
    [1,0,0,1],
    [1,0,0,1]
  ],
  'P': [
    [1,1,1,0],
    [1,0,0,1],
    [1,1,1,0],
    [1,0,0,0],
    [1,0,0,0]
  ]
};

function renderIconPNG(size) {
  const width = size;
  const height = size;
  const buffer = Buffer.alloc(width * height * 4);

  const cx = width / 2;
  const cy = height / 2;
  const scale = size / 512;

  // 1. Draw Rounded Gradient Background
  const rx = 112 * scale;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const idx = (y * width + x) * 4;
      
      // Check rounded corner clip
      let inBounds = true;
      if (x < rx && y < rx) {
        if (Math.hypot(x - rx, y - rx) > rx) inBounds = false;
      } else if (x > width - rx && y < rx) {
        if (Math.hypot(x - (width - rx), y - rx) > rx) inBounds = false;
      } else if (x < rx && y > height - rx) {
        if (Math.hypot(x - rx, y - (height - rx)) > rx) inBounds = false;
      } else if (x > width - rx && y > height - rx) {
        if (Math.hypot(x - (width - rx), y - (height - rx)) > rx) inBounds = false;
      }

      if (!inBounds) {
        buffer[idx] = 0; buffer[idx+1] = 0; buffer[idx+2] = 0; buffer[idx+3] = 0;
        continue;
      }

      // Gradient indigo: #4f46e5 (79, 70, 229) to #312e81 (49, 46, 129)
      const factor = y / height;
      buffer[idx] = Math.round(79 * (1 - factor) + 49 * factor);     // R
      buffer[idx + 1] = Math.round(70 * (1 - factor) + 46 * factor); // G
      buffer[idx + 2] = Math.round(229 * (1 - factor) + 129 * factor);// B
      buffer[idx + 3] = 255;                                          // Alpha
    }
  }

  // Helper to draw filled circle
  function drawCircle(centerX, centerY, radius, strokeWidth, r, g, b, a) {
    for (let y = Math.floor(centerY - radius - strokeWidth); y <= Math.ceil(centerY + radius + strokeWidth); y++) {
      for (let x = Math.floor(centerX - radius - strokeWidth); x <= Math.ceil(centerX + radius + strokeWidth); x++) {
        if (x < 0 || x >= width || y < 0 || y >= height) continue;
        const dist = Math.hypot(x - centerX, y - centerY);
        if (Math.abs(dist - radius) <= strokeWidth / 2) {
          const idx = (y * width + x) * 4;
          buffer[idx] = r; buffer[idx+1] = g; buffer[idx+2] = b; buffer[idx+3] = a;
        }
      }
    }
  }

  // Draw Circle ring around badge
  drawCircle(cx, cy - 20 * scale, 135 * scale, 10 * scale, 255, 255, 255, 70);

  // Draw Shield Badge (fill)
  const shieldTop = cy - 100 * scale;
  const shieldBottom = cy + 120 * scale;
  const shieldWidth = 140 * scale;
  for (let y = Math.floor(shieldTop); y <= Math.ceil(shieldBottom); y++) {
    const progress = (y - shieldTop) / (shieldBottom - shieldTop);
    let curHalfWidth = shieldWidth;
    if (progress > 0.4) {
      const p = (progress - 0.4) / 0.6;
      curHalfWidth = shieldWidth * (1 - p * p);
    }
    for (let x = Math.floor(cx - curHalfWidth); x <= Math.ceil(cx + curHalfWidth); x++) {
      if (x < 0 || x >= width || y < 0 || y >= height) continue;
      const idx = (y * width + x) * 4;
      // White shield badge opacity
      buffer[idx] = 255; buffer[idx+1] = 255; buffer[idx+2] = 255; buffer[idx+3] = 245;
    }
  }

  // Draw Inner Indigo Shield
  const innerTop = shieldTop + 20 * scale;
  const innerBottom = shieldBottom - 25 * scale;
  const innerWidth = shieldWidth - 20 * scale;
  for (let y = Math.floor(innerTop); y <= Math.ceil(innerBottom); y++) {
    const progress = (y - innerTop) / (innerBottom - innerTop);
    let curHalfWidth = innerWidth;
    if (progress > 0.4) {
      const p = (progress - 0.4) / 0.6;
      curHalfWidth = innerWidth * (1 - p * p);
    }
    for (let x = Math.floor(cx - curHalfWidth); x <= Math.ceil(cx + curHalfWidth); x++) {
      if (x < 0 || x >= width || y < 0 || y >= height) continue;
      const idx = (y * width + x) * 4;
      buffer[idx] = 79; buffer[idx+1] = 70; buffer[idx+2] = 229; buffer[idx+3] = 255;
    }
  }

  // Helper to draw text using bitmap font
  function drawText(text, startX, startY, pixelSize) {
    let cursorX = startX;
    for (let char of text) {
      if (char === ' ') {
        cursorX += pixelSize * 3;
        continue;
      }
      const grid = fontMap[char];
      if (!grid) continue;
      for (let r = 0; r < grid.length; r++) {
        for (let c = 0; c < grid[r].length; c++) {
          if (grid[r][c] === 1) {
            for (let py = 0; py < pixelSize; py++) {
              for (let px = 0; px < pixelSize; px++) {
                const drawX = Math.floor(cursorX + c * pixelSize + px);
                const drawY = Math.floor(startY + r * pixelSize + py);
                if (drawX >= 0 && drawX < width && drawY >= 0 && drawY < height) {
                  const idx = (drawY * width + drawX) * 4;
                  buffer[idx] = 255; buffer[idx+1] = 255; buffer[idx+2] = 255; buffer[idx+3] = 255;
                }
              }
            }
          }
        }
      }
      cursorX += (grid[0].length + 1) * pixelSize;
    }
  }

  // Render "A I M S" text inside shield
  const aimsPixelSize = Math.max(3, Math.round(10 * scale));
  const aimsWidth = (4*4 + 3*3 + 5*4 + 4*4 + 3*4) * aimsPixelSize; // rough width
  drawText("A I M S", cx - 95 * scale, cy - 22 * scale, aimsPixelSize);

  // Render "H R I S" text at bottom
  const hrisPixelSize = Math.max(2, Math.round(7 * scale));
  drawText("H R I S", cx - 75 * scale, cy + 155 * scale, hrisPixelSize);

  return encodePNG(buffer, width, height);
}

function encodePNG(buffer, width, height) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  
  const ihdrChunk = createChunk('IHDR', ihdr);
  
  const rawData = Buffer.alloc(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) {
    rawData[y * (width * 4 + 1)] = 0; // Filter type 0
    buffer.copy(rawData, y * (width * 4 + 1) + 1, y * width * 4, (y + 1) * width * 4);
  }
  const compressed = zlib.deflateSync(rawData);
  const idatChunk = createChunk('IDAT', compressed);
  const iendChunk = createChunk('IEND', Buffer.alloc(0));
  
  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

function createChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const body = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([len, body, crc]);
}

function crc32(buf) {
  let c = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c ^= buf[n];
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
  }
  return (c ^ 0xffffffff) >>> 0;
}

fs.writeFileSync('icon-192.png', renderIconPNG(192));
fs.writeFileSync('icon-512.png', renderIconPNG(512));
fs.writeFileSync('apple-touch-icon.png', renderIconPNG(180));
console.log('High Quality AIMS PNG Icons Generated!');
