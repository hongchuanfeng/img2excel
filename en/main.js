const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const previewContainer = document.getElementById('previewContainer');
const previewImage = document.getElementById('previewImage');
const pixelGrid = document.getElementById('pixelGrid');
const canvas = document.getElementById('canvas');
const generateButton = document.getElementById('generateButton');
const loadingText = document.getElementById('loading');
const settings = document.getElementById('settings');

// 获取所有设置输入
const maxSizeInput = document.getElementById('maxSize');
const brightnessInput = document.getElementById('brightness');
const contrastInput = document.getElementById('contrast');
const saturationInput = document.getElementById('saturation');
const hueInput = document.getElementById('hue');

// Excel设置默认值
const cellSize = 40;
const sheetBackground = '#FFFFFF';
const borderStyle = 'none';
const borderColor = '#000000';
const zoomScale = 100;

let pixelData = [];
let originalImageData = null;

// 拖拽上传
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith('image/')) {
        processImage(file);
    }
});

// 点击上传
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        processImage(file);
    }
});

// 图片处理函数
function processImage(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
            originalImageData = {
                img: img,
                src: e.target.result
            };
            applyImageAdjustments();
        };
        img.src = e.target.result;
    };
    reader.readAsDataURL(file);
}

// 应用图片调整
function applyImageAdjustments() {
    if (!originalImageData) return;

    const maxSize = parseInt(maxSizeInput.value);
    const ctx = canvas.getContext('2d');

    // 计算缩放比例
    const scale = Math.min(maxSize / originalImageData.img.width, maxSize / originalImageData.img.height);
    const width = Math.floor(originalImageData.img.width * scale);
    const height = Math.floor(originalImageData.img.height * scale);

    canvas.width = width;
    canvas.height = height;

    // 应用滤镜
    ctx.filter = `
                brightness(${100 + parseInt(brightnessInput.value)}%)
                contrast(${100 + parseInt(contrastInput.value)}%)
                saturate(${100 + parseInt(saturationInput.value)}%)
                hue-rotate(${parseInt(hueInput.value)}deg)
            `;

    // 绘制并获取像素数据
    ctx.drawImage(originalImageData.img, 0, 0, width, height);
    const imageData = ctx.getImageData(0, 0, width, height);

    // 转换为RGB数组
    pixelData = [];
    for (let i = 0; i < height; i++) {
        const row = [];
        for (let j = 0; j < width; j++) {
            const index = (i * width + j) * 4;
            row.push({
                r: imageData.data[index],
                g: imageData.data[index + 1],
                b: imageData.data[index + 2]
            });
        }
        pixelData.push(row);
    }

    // 显示预览
    previewImage.src = canvas.toDataURL();
    previewImage.style.display = 'block';
    document.getElementById('noImageMessage').style.display = 'none';
    document.querySelector('.modern-pixel-preview-card').style.display = 'block';
    generateButton.style.display = 'block';

    // 更新像素预览
    updatePixelPreview();
}

// 监听图片调整参数变化
brightnessInput.addEventListener('input', (e) => {
    const valueSpan = e.target.parentElement.querySelector('.range-value');
    if (valueSpan) valueSpan.textContent = e.target.value;
    applyImageAdjustments();
});
contrastInput.addEventListener('input', (e) => {
    const valueSpan = e.target.parentElement.querySelector('.range-value');
    if (valueSpan) valueSpan.textContent = e.target.value;
    applyImageAdjustments();
});
saturationInput.addEventListener('input', (e) => {
    const valueSpan = e.target.parentElement.querySelector('.range-value');
    if (valueSpan) valueSpan.textContent = e.target.value;
    applyImageAdjustments();
});
hueInput.addEventListener('input', (e) => {
    const valueSpan = e.target.parentElement.querySelector('.range-value');
    if (valueSpan) valueSpan.textContent = e.target.value;
    applyImageAdjustments();
});
maxSizeInput.addEventListener('change', applyImageAdjustments);

// 更新像素预览
function updatePixelPreview() {
    pixelGrid.style.gridTemplateColumns = `repeat(${pixelData[0].length}, 10px)`;
    pixelGrid.innerHTML = '';

    pixelData.forEach(row => {
        row.forEach(pixel => {
            const div = document.createElement('div');
            div.className = 'pixel';
            div.style.backgroundColor = `rgb(${pixel.r},${pixel.g},${pixel.b})`;
            pixelGrid.appendChild(div);
        });
    });
}

generateButton.addEventListener('click', async () => {
    document.getElementById('loading').style.display = 'block';
    generateButton.disabled = true;

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Pixel Art');

        // Excel中列宽单位约为8像素，行高单位为像素
        const columnWidth = cellSize / 6;

        // 设置列宽
        pixelData[0].forEach((_, index) => {
            const col = worksheet.getColumn(index + 1);
            col.width = columnWidth;
        });

        // 设置行高
        pixelData.forEach((_, index) => {
            const row = worksheet.getRow(index + 1);
            row.height = cellSize;
        });

        // 设置工作表背景色
        worksheet.properties.tabColor = {
            argb: 'FF' + sheetBackground.substring(1)
        };

        // 填充颜色
        pixelData.forEach((row, rowIndex) => {
            row.forEach((pixel, colIndex) => {
                const cell = worksheet.getCell(rowIndex + 1, colIndex + 1);

                cell.value = '';
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: `FF${toHex(pixel.r)}${toHex(pixel.g)}${toHex(pixel.b)}` }
                };

                // 设置边框
                if (borderStyle !== 'none') {
                    cell.border = {
                        top: { style: borderStyle, color: { argb: 'FF' + borderColor.substring(1) } },
                        left: { style: borderStyle, color: { argb: 'FF' + borderColor.substring(1) } },
                        bottom: { style: borderStyle, color: { argb: 'FF' + borderColor.substring(1) } },
                        right: { style: borderStyle, color: { argb: 'FF' + borderColor.substring(1) } }
                    };
                }
            });
        });

        // 工作表视图设置
        worksheet.views = [
            {
                showGridLines: false,
                zoomScale: zoomScale
            }
        ];

        // 生成并下载文件
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'pixel_art.xlsx';
        link.click();
    } catch (error) {
        console.error('生成Excel文件时发生错误:', error);
        alert('生成Excel文件时发生错误，请重试');
    } finally {
        document.getElementById('loading').style.display = 'none';
        generateButton.disabled = false;
    }
});

// RGB转十六进制辅助函数
function toHex(num) {
    const hex = num.toString(16).toUpperCase();
    return hex.length === 1 ? '0' + hex : hex;
}