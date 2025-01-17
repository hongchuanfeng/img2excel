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
const cellSizeInput = document.getElementById('cellSize');
const brightnessInput = document.getElementById('brightness');
const contrastInput = document.getElementById('contrast');
const saturationInput = document.getElementById('saturation');
const hueInput = document.getElementById('hue');
const sheetBackgroundInput = document.getElementById('sheetBackground');
const borderStyleInput = document.getElementById('borderStyle');
const borderColorInput = document.getElementById('borderColor');
const zoomScaleInput = document.getElementById('zoomScale');

let pixelData = [];
let originalImageData = null;

// 拖拽上传
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#45a049';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.borderColor = '#ccc';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#ccc';
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
    previewContainer.style.display = 'block';
    settings.style.display = 'block';

    // 更新像素预览
    updatePixelPreview();
}

// 监听图片调整参数变化
brightnessInput.addEventListener('input', applyImageAdjustments);
contrastInput.addEventListener('input', applyImageAdjustments);
saturationInput.addEventListener('input', applyImageAdjustments);
hueInput.addEventListener('input', applyImageAdjustments);
maxSizeInput.addEventListener('change', applyImageAdjustments);

// 更新像素预览
function updatePixelPreview() {
    pixelGrid.style.gridTemplateColumns = `repeat(${pixelData[0].length}, 8px)`;
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
    loadingText.style.display = 'block';
    generateButton.disabled = true;

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Pixel Art');
        const cellSize = parseInt(cellSizeInput.value);

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
            argb: 'FF' + sheetBackgroundInput.value.substring(1)
        };

        // 获取边框样式
        const borderStyle = borderStyleInput.value;
        const borderColor = borderColorInput.value;

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
                zoomScale: parseInt(zoomScaleInput.value)
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
        loadingText.style.display = 'none';
        generateButton.disabled = false;
    }
});

// RGB转十六进制辅助函数
function toHex(num) {
    const hex = num.toString(16).toUpperCase();
    return hex.length === 1 ? '0' + hex : hex;
}