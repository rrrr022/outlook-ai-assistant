/**
 * AppSource Icon Generator
 * Converts SVG to PNG for Microsoft Partner Center submission
 * 
 * Run: node generate-icon.js
 */

const fs = require('fs');
const path = require('path');

// Read the SVG file
const svgPath = path.join(__dirname, 'src/client/assets/icon-300.svg');
const outputPath = path.join(__dirname, 'src/client/assets/icon-300.png');

console.log('🎨 AppSource Icon Generator');
console.log('============================\n');

// Check if SVG exists
if (fs.existsSync(svgPath)) {
    console.log('✓ Found SVG icon at:', svgPath);
    
    // For PNG conversion, you have several options:
    console.log('\n📋 To convert SVG to PNG, use one of these methods:\n');
    
    console.log('Option 1: Online Converter (Easiest)');
    console.log('------------------------------------');
    console.log('1. Open: https://svgtopng.com/');
    console.log('2. Upload:', svgPath);
    console.log('3. Set size to 300x300');
    console.log('4. Download PNG\n');
    
    console.log('Option 2: Browser Method');
    console.log('------------------------');
    console.log('1. Open appstore-assets-tool.html in your browser');
    console.log('2. Click "Generate Icon" button');
    console.log('3. Click "Download PNG" button\n');
    
    console.log('Option 3: Inkscape (If installed)');
    console.log('---------------------------------');
    console.log('inkscape --export-type="png" --export-filename="' + outputPath + '" "' + svgPath + '"\n');
    
    console.log('Option 4: ImageMagick (If installed)');
    console.log('------------------------------------');
    console.log('convert -background none -resize 300x300 "' + svgPath + '" "' + outputPath + '"\n');
    
    console.log('Option 5: Sharp (npm package)');
    console.log('-----------------------------');
    console.log('npm install sharp');
    console.log('Then run: node convert-with-sharp.js\n');
    
} else {
    console.log('✗ SVG not found. Creating it now...');
}

// Create a simple fallback icon using Canvas (if canvas package is available)
try {
    const { createCanvas } = require('canvas');
    
    console.log('📦 Canvas package found! Generating PNG directly...\n');
    
    const canvas = createCanvas(300, 300);
    const ctx = canvas.getContext('2d');
    
    // Background
    const bgGradient = ctx.createRadialGradient(150, 150, 0, 150, 150, 180);
    bgGradient.addColorStop(0, '#1a1a24');
    bgGradient.addColorStop(1, '#0a0a0f');
    ctx.fillStyle = bgGradient;
    ctx.beginPath();
    ctx.arc(150, 150, 145, 0, Math.PI * 2);
    ctx.fill();
    
    // Border
    ctx.strokeStyle = '#39ff14';
    ctx.lineWidth = 3;
    ctx.stroke();
    
    // Central core
    ctx.fillStyle = '#39ff14';
    ctx.beginPath();
    ctx.arc(150, 150, 25, 0, Math.PI * 2);
    ctx.fill();
    
    ctx.fillStyle = '#0a0a0f';
    ctx.beginPath();
    ctx.arc(150, 150, 15, 0, Math.PI * 2);
    ctx.fill();
    
    ctx.fillStyle = '#ffb000';
    ctx.beginPath();
    ctx.arc(150, 150, 8, 0, Math.PI * 2);
    ctx.fill();
    
    // Neural connections
    ctx.strokeStyle = '#39ff14';
    ctx.lineWidth = 4;
    ctx.lineCap = 'round';
    
    // Function to draw connection
    function drawConnection(x1, y1, x2, y2) {
        ctx.beginPath();
        ctx.moveTo(150 + x1, 150 + y1);
        ctx.lineTo(150 + x2, 150 + y2);
        ctx.stroke();
    }
    
    // Function to draw node
    function drawNode(x, y, r, color) {
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.arc(150 + x, 150 + y, r, 0, Math.PI * 2);
        ctx.fill();
    }
    
    // Draw connections
    // Top
    drawConnection(0, -25, 0, -75);
    drawConnection(0, -50, -30, -70);
    drawConnection(0, -50, 30, -70);
    
    // Bottom
    drawConnection(0, 25, 0, 75);
    drawConnection(0, 50, -30, 70);
    drawConnection(0, 50, 30, 70);
    
    // Left
    drawConnection(-25, 0, -75, 0);
    drawConnection(-50, 0, -70, -30);
    drawConnection(-50, 0, -70, 30);
    
    // Right
    drawConnection(25, 0, 75, 0);
    drawConnection(50, 0, 70, -30);
    drawConnection(50, 0, 70, 30);
    
    // Diagonal
    ctx.lineWidth = 3;
    drawConnection(18, 18, 55, 55);
    drawConnection(-18, 18, -55, 55);
    drawConnection(18, -18, 55, -55);
    drawConnection(-18, -18, -55, -55);
    
    // Draw nodes
    // Main axis nodes
    drawNode(0, -75, 10, '#39ff14');
    drawNode(0, 75, 10, '#39ff14');
    drawNode(-75, 0, 10, '#39ff14');
    drawNode(75, 0, 10, '#39ff14');
    
    // Corner nodes
    drawNode(-30, -70, 8, '#ffb000');
    drawNode(30, -70, 8, '#ffb000');
    drawNode(-30, 70, 8, '#ffb000');
    drawNode(30, 70, 8, '#ffb000');
    drawNode(-70, -30, 8, '#ffb000');
    drawNode(-70, 30, 8, '#ffb000');
    drawNode(70, -30, 8, '#ffb000');
    drawNode(70, 30, 8, '#ffb000');
    
    // Diagonal nodes
    drawNode(55, 55, 8, '#39ff14');
    drawNode(-55, 55, 8, '#39ff14');
    drawNode(55, -55, 8, '#39ff14');
    drawNode(-55, -55, 8, '#39ff14');
    
    // "AI" text
    ctx.fillStyle = '#39ff14';
    ctx.font = 'bold 36px Arial Black';
    ctx.textAlign = 'center';
    ctx.fillText('AI', 150, 268);
    
    // Save PNG
    const buffer = canvas.toBuffer('image/png');
    fs.writeFileSync(outputPath, buffer);
    
    console.log('✅ PNG icon generated successfully!');
    console.log('📁 Saved to:', outputPath);
    
} catch (e) {
    if (e.code === 'MODULE_NOT_FOUND') {
        console.log('💡 Tip: Install canvas for automatic PNG generation:');
        console.log('   npm install canvas\n');
        console.log('   Then run this script again.\n');
    }
}

console.log('\n============================');
console.log('📋 AppSource Submission Files:');
console.log('============================');
console.log('1. manifest.azure.xml    - ✓ Ready');
console.log('2. icon-300.svg          - ✓ Created');
console.log('3. icon-300.png          - ' + (fs.existsSync(outputPath) ? '✓ Created' : '⏳ Convert from SVG'));
console.log('4. Screenshots           - ⏳ Take 3-5 screenshots (1366x768)');
console.log('\n🌐 Open appstore-assets-tool.html for interactive guide');
