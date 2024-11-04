const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const ExcelJS = require("exceljs");
const readline = require("readline");

const run = () => {
  // Add these console styling functions at the top
  const consoleStyles = {
    success: "\x1b[92m",
    info: "\x1b[96m",
    warning: "\x1b[93m",
    error: "\x1b[91m",
    reset: "\x1b[0m",
  };

  function logSuccess(message) {
    console.log(`${consoleStyles.success}✓ ${message}${consoleStyles.reset}`);
  }

  function logInfo(message) {
    console.log(`${consoleStyles.info}ℹ ${message}${consoleStyles.reset}`);
  }

  function logWarning(message) {
    console.log(`${consoleStyles.warning}⚠ ${message}${consoleStyles.reset}`);
  }

  function logError(message) {
    console.error(`${consoleStyles.error}✗ ${message}${consoleStyles.reset}`);
  }

  // Update the directory paths to use absolute paths from executable location
  const appRoot = process.cwd();
  const resultDir = path.join(appRoot, "result");
  const imagesDir = path.join(appRoot, "images");
  const textDir = path.join(appRoot, "text");

  // Function to check required directories
  function checkRequiredDirectories() {
    console.log(
      `${consoleStyles.info}Checking required directories...${consoleStyles.reset}\n`
    );

    const required = [
      { path: resultDir, name: "result" },
      { path: imagesDir, name: "images" },
      { path: textDir, name: "text" },
    ];

    let createdDirs = [];
    for (const dir of required) {
      process.stdout.write(`Checking ${dir.name} directory... `);
      if (!fs.existsSync(dir.path)) {
        try {
          fs.mkdirSync(dir.path, { recursive: true });
          createdDirs.push(dir.name);
          logSuccess("Created");
        } catch (error) {
          logError(`Failed to create ${dir.name} directory: ${error.message}`);
          process.exit(1);
        }
      } else {
        logSuccess("Found");
      }
    }

    if (createdDirs.length > 0) {
      console.log(
        `\n${
          consoleStyles.warning
        }Note: Please add required files to: ${createdDirs.join(", ")}${
          consoleStyles.reset
        }`
      );
    }

    // Helper function to check files recursively in directory
    function hasFilesRecursive(dir, validExtensions) {
      const items = fs.readdirSync(dir);

      for (const item of items) {
        const fullPath = path.join(dir, item);
        const stat = fs.statSync(fullPath);

        if (stat.isDirectory()) {
          // Recursively check subdirectories
          if (hasFilesRecursive(fullPath, validExtensions)) {
            return true;
          }
        } else {
          const ext = path.extname(item).toLowerCase();
          if (validExtensions.includes(ext)) {
            return true;
          }
        }
      }

      return false;
    }

    // Check if essential directories have files (including subfolders)
    const hasImages = hasFilesRecursive(imagesDir, [
      ".jpg",
      ".jpeg",
      ".png",
      ".gif",
      ".webp",
    ]);
    const hasTextFiles = hasFilesRecursive(textDir, [".xlsx", ".xls"]);

    if (!hasImages || !hasTextFiles) {
      console.log("\nMissing required files:");
      if (!hasImages)
        logWarning("No images found in images directory or its subfolders");
      if (!hasTextFiles)
        logWarning("No Excel files found in text directory or its subfolders");
      logInfo("Please add the required files and run the program again.");
      process.exit(0);
    }
  }

  // Function to get all images from images folder
  function getAllImages(dir = imagesDir) {
    if (!fs.existsSync(dir)) {
      console.error(
        "Images directory does not exist. Please create it and add images."
      );
      return [];
    }

    const images = [];
    const items = fs.readdirSync(dir);

    for (const item of items) {
      const fullPath = path.join(dir, item);
      const stat = fs.statSync(fullPath);

      if (stat.isDirectory()) {
        // Recursively get images from subdirectories
        const subImages = getAllImages(fullPath);
        images.push(...subImages);
      } else {
        const ext = path.extname(item).toLowerCase();
        if ([".jpg", ".jpeg", ".png", ".gif", ".webp"].includes(ext)) {
          // Store relative path from images directory
          images.push(path.relative(imagesDir, fullPath));
        }
      }
    }

    return images;
  }

  // Add efficient caching with size limits and auto-cleanup
  const imageCache = new Map();
  const MAX_CACHE_SIZE = 100 * 1024 * 1024; // 100MB cache limit
  let currentCacheSize = 0;

  function clearOldestCache() {
    if (imageCache.size === 0) return;
    const oldestKey = imageCache.keys().next().value;
    const oldestSize = imageCache.get(oldestKey).length;
    imageCache.delete(oldestKey);
    currentCacheSize -= oldestSize;
  }

  // Optimized image cache management
  function readImageFile(imagePath) {
    try {
      if (imageCache.has(imagePath)) {
        return imageCache.get(imagePath);
      }

      const fullPath = path.join(imagesDir, imagePath);
      const imageBuffer = fs.readFileSync(fullPath);

      // Manage cache size
      while (currentCacheSize + imageBuffer.length > MAX_CACHE_SIZE) {
        clearOldestCache();
      }

      imageCache.set(imagePath, imageBuffer);
      currentCacheSize += imageBuffer.length;
      return imageBuffer;
    } catch (error) {
      console.error("Error reading image:", error);
      throw error;
    }
  }

  // Function to invert image colors
  async function invertImageColors(imageBuffer) {
    try {
      const invertedImage = await sharp(imageBuffer).negate().toBuffer();
      return invertedImage;
    } catch (error) {
      console.error("Error inverting image:", error);
      throw error;
    }
  }

  // Optimize SVG generation
  function createSVGBuffer(width, height, elements) {
    return Buffer.from(`
        <svg width="${width}" height="${height}">
            ${elements.join("")}
        </svg>
    `);
  }

  // Optimized drawing functions
  async function drawEdgePoints(imageBuffer, edgePoints, color = "red") {
    try {
      const metadata = await sharp(imageBuffer).metadata();
      const svgElements = edgePoints.map(
        (point) =>
          `<circle cx="${point.x}" cy="${point.y}" r="1" fill="${color}"/>`
      );

      const svgBuffer = createSVGBuffer(
        metadata.width,
        metadata.height,
        svgElements
      );
      return sharp(imageBuffer)
        .composite([{ input: svgBuffer, top: 0, left: 0 }])
        .toBuffer();
    } catch (error) {
      console.error("Error drawing edge points:", error);
      throw error;
    }
  }

  // Function to draw shapes on image
  async function drawObjects(imageBuffer, objects) {
    try {
      const image = sharp(imageBuffer);
      const metadata = await image.metadata();

      // Create SVG buffer for drawing
      const svgBuffer = Buffer.from(`
            <svg width="${metadata.width}" height="${metadata.height}">
                ${objects
                  .map((obj) => {
                    switch (obj.type) {
                      case "rectangle":
                        return `<rect 
                                x="${obj.x}" 
                                y="${obj.y}" 
                                width="${obj.width}" 
                                height="${obj.height}" 
                                fill="none" 
                                stroke="${obj.color || "red"}" 
                                stroke-width="${obj.strokeWidth || 2}"
                            />`;
                      case "circle":
                        return `<circle 
                                cx="${obj.x}" 
                                cy="${obj.y}" 
                                r="${obj.radius}" 
                                fill="none" 
                                stroke="${obj.color || "red"}" 
                                stroke-width="${obj.strokeWidth || 2}"
                            />`;
                      case "line":
                        return `<line 
                                x1="${obj.x1}" 
                                y1="${obj.y1}" 
                                x2="${obj.x2}" 
                                y2="${obj.y2}" 
                                stroke="${obj.color || "red"}" 
                                stroke-width="${obj.strokeWidth || 2}"
                            />`;
                      case "text":
                        return `<text 
                                x="${obj.x}" 
                                y="${obj.y}" 
                                fill="${obj.color || "red"}" 
                                font-size="${obj.fontSize || 20}"
                                text-anchor="${obj.textAnchor || "start"}"
                                dominant-baseline="${
                                  obj.dominantBaseline || "auto"
                                }"
                                font-weight="${obj.fontWeight || "normal"}"
                                font-family="${obj.fontFamily || "Arial"}"
                                style="${obj.style || ""}"
                            >${obj.text}</text>`;
                      default:
                        return "";
                    }
                  })
                  .join("")}
            </svg>
        `);

      const result = await image
        .composite([
          {
            input: svgBuffer,
            top: 0,
            left: 0,
          },
        ])
        .toBuffer();

      return result;
    } catch (error) {
      console.error("Error drawing objects:", error);
      throw error;
    }
  }

  // Function to find empty regions in the image
  async function findEmptyRegions(imageBuffer, minWidth = 100, minHeight = 30) {
    try {
      // Convert to grayscale and get raw pixels
      const { data, info } = await sharp(imageBuffer)
        .grayscale()
        .raw()
        .toBuffer({ resolveWithObject: true });

      const width = info.width;
      const height = info.height;
      const regions = [];
      const cellSize = 5; // Smaller cell size for more precise detection

      // Create grid to track empty areas
      const gridWidth = Math.ceil(width / cellSize);
      const gridHeight = Math.ceil(height / cellSize);
      const grid = new Uint8Array(gridWidth * gridHeight);

      // Mark cells as occupied based on pixel intensity
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          const pixel = data[y * width + x];
          // If pixel is not close to white (threshold 200), mark cell as occupied
          if (pixel < 200) {
            const gridX = Math.floor(x / cellSize);
            const gridY = Math.floor(y / cellSize);
            grid[gridY * gridWidth + gridX] = 1;
          }
        }
      }

      // Find full-width empty regions
      let currentEmptyHeight = 0;
      let startY = 0;

      for (let y = 0; y < gridHeight; y++) {
        // Check if current row is empty
        let isRowEmpty = true;
        for (let x = 0; x < gridWidth; x++) {
          if (grid[y * gridWidth + x] === 1) {
            isRowEmpty = false;
            break;
          }
        }

        if (isRowEmpty) {
          if (currentEmptyHeight === 0) {
            startY = y;
          }
          currentEmptyHeight++;
        } else {
          if (currentEmptyHeight >= Math.ceil(minHeight / cellSize)) {
            regions.push({
              x: 0,
              y: startY * cellSize,
              width: width,
              height: currentEmptyHeight * cellSize,
            });
          }
          currentEmptyHeight = 0;
        }
      }

      // Check last region if image ends with empty space
      if (currentEmptyHeight >= Math.ceil(minHeight / cellSize)) {
        regions.push({
          x: 0,
          y: startY * cellSize,
          width: width,
          height: currentEmptyHeight * cellSize,
        });
      }

      // Filter and merge close regions
      const mergedRegions = [];
      let currentRegion = null;

      for (const region of regions) {
        if (!currentRegion) {
          currentRegion = region;
          continue;
        }

        // If regions are close, merge them
        if (
          region.y - (currentRegion.y + currentRegion.height) <
          cellSize * 4
        ) {
          currentRegion.height = region.y + region.height - currentRegion.y;
        } else {
          mergedRegions.push(currentRegion);
          currentRegion = region;
        }
      }

      if (currentRegion) {
        mergedRegions.push(currentRegion);
      }

      // Convert to rectangle objects
      return mergedRegions
        .filter((region) => region.height >= minHeight)
        .map((region) => ({
          type: "rectangle",
          ...region,
          color: "blue",
          strokeWidth: 1,
          area: region.width * region.height, // Add area property for sorting
        }))
        .sort((a, b) => b.area - a.area); // Sort by area in descending order
    } catch (error) {
      console.error("Error finding empty regions:", error);
      throw error;
    }
  }

  // Helper function to get format properties for a text position
  function getFormatProperties(position, colorSegments, boldSegments) {
    const format = {
      color: "#000000",
      isBold: false,
    };

    // Check color segments
    const colorSegment = colorSegments.find(
      (seg) => position >= seg.start && position < seg.start + seg.length
    );
    if (colorSegment) {
      format.color = colorSegment.color;
    }

    // Check bold segments
    const boldSegment = boldSegments.find(
      (seg) => position >= seg.start && position < seg.start + seg.length
    );
    if (boldSegment) {
      format.isBold = true;
    }

    return format;
  }

  // Update wrapText to handle formatting
  function wrapText(
    text,
    maxWidth,
    fontSize = 30,
    x = 0,
    y = 0,
    colorSegments = [],
    boldSegments = []
  ) {
    if (!text) return [];

    // Preallocate arrays
    const lines = new Array(Math.ceil(text.length / 50));
    const words = text.split(/\s+/);
    let lineIndex = 0;
    let textPosition = 0;

    let currentLine = {
      x,
      y,
      segments: [],
      width: 0,
      lineHeight: fontSize * 1.3,
    };

    // Process words efficiently with formatting
    for (let i = 0; i < words.length; i++) {
      const word = words[i];
      const wordStart = text.indexOf(word, textPosition);
      const format = getFormatProperties(
        wordStart,
        colorSegments,
        boldSegments
      );

      // Calculate width based on bold status
      const wordWidth = calculateTextWidth(
        word,
        format.isBold ? fontSize * 1.1 : fontSize
      );
      const spaceWidth = i > 0 ? calculateTextWidth(" ", fontSize) : 0;

      if (
        currentLine.width + wordWidth + spaceWidth > maxWidth &&
        currentLine.segments.length > 0
      ) {
        // Complete current line
        lines[lineIndex++] = currentLine;

        // Start new line
        currentLine = {
          x,
          y: y + lineIndex * fontSize * 1.3,
          segments: [
            {
              text: word,
              width: wordWidth,
              color: format.color,
              isBold: format.isBold,
            },
          ],
          width: wordWidth,
          lineHeight: fontSize * 1.3,
        };
      } else {
        // Add to current line
        if (currentLine.segments.length > 0) {
          currentLine.width += spaceWidth;
        }
        currentLine.segments.push({
          text: word,
          width: wordWidth,
          color: format.color,
          isBold: format.isBold,
        });
        currentLine.width += wordWidth;
      }

      textPosition = wordStart + word.length;
    }

    // Add final line if not empty
    if (currentLine.segments.length > 0) {
      lines[lineIndex++] = currentLine;
    }

    return lines.slice(0, lineIndex);
  }

  // Update drawRegionsWithText function to remove debug logging
  async function drawRegionsWithText(imageBuffer, regions, textObjects) {
    try {
      const metadata = await sharp(imageBuffer).metadata();

      // Calculate upscale factor based on image dimensions
      const UPSCALE_THRESHOLD = 1200; // Minimum dimension for upscaling
      const MAX_UPSCALE = 2; // Maximum upscale factor

      let upscaleFactor = 1;
      if (Math.min(metadata.width, metadata.height) < UPSCALE_THRESHOLD) {
        upscaleFactor = Math.min(
          MAX_UPSCALE,
          UPSCALE_THRESHOLD / Math.min(metadata.width, metadata.height)
        );
      }

      // Upscale dimensions
      const targetWidth = Math.round(metadata.width * upscaleFactor);
      const targetHeight = Math.round(metadata.height * upscaleFactor);

      // Upscale image with quality settings
      const upscaledImage = await sharp(imageBuffer)
        .resize(targetWidth, targetHeight, {
          kernel: "lanczos3",
          fit: "fill",
          withoutEnlargement: false,
        })
        .toBuffer();

      // Scale text object coordinates
      const scaledTextObjects = textObjects.map((obj) => ({
        ...obj,
        x: obj.x * upscaleFactor,
        y: obj.y * upscaleFactor,
        fontSize: obj.fontSize * upscaleFactor,
      }));

      const svgElements = [];

      // Add text elements with proper SVG escaping and rendering settings
      scaledTextObjects.forEach((obj) => {
        if (obj.colorSegments && obj.colorSegments.length > 0) {
          const textElement = `
          <text 
            x="${obj.x}" 
            y="${obj.y}"
            font-size="${obj.fontSize}px"
            text-anchor="${obj.textAnchor}"
            dominant-baseline="${obj.dominantBaseline}"
            font-weight="${obj.fontWeight}"
            font-family="${obj.fontFamily}"
            style="${obj.style}"
            shape-rendering="geometricPrecision"
            text-rendering="geometricPrecision"
          >
            ${obj.colorSegments
              .map((segment, index) => {
                const escapedText = segment.text
                  .replace(/&/g, "&amp;")
                  .replace(/</g, "&lt;")
                  .replace(/>/g, "&gt;")
                  .replace(/"/g, "&quot;")
                  .replace(/'/g, "&apos;");

                return `<tspan 
                fill="${segment.color}"
                ${
                  index > 0
                    ? `dx="${calculateTextWidth(" ", obj.fontSize)}"`
                    : ""
                }
                shape-rendering="geometricPrecision"
                text-rendering="geometricPrecision"
              >${escapedText}</tspan>`;
              })
              .join("")}
          </text>`;

          svgElements.push(textElement);
        } else {
          const escapedText = obj.text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");

          svgElements.push(`
          <text 
            x="${obj.x}" 
            y="${obj.y}"
            fill="#000000"
            font-size="${obj.fontSize}px"
            text-anchor="${obj.textAnchor}"
            dominant-baseline="${obj.dominantBaseline}"
            font-weight="${obj.fontWeight}"
            font-family="${obj.fontFamily}"
            style="${obj.style}"
            shape-rendering="geometricPrecision"
            text-rendering="geometricPrecision"
          >${escapedText}</text>
        `);
        }
      });

      // Create SVG with viewBox to maintain proper scaling
      const svgBuffer = Buffer.from(`<?xml version="1.0" encoding="UTF-8"?>
      <svg 
        xmlns="http://www.w3.org/2000/svg"
        width="${targetWidth}" 
        height="${targetHeight}"
        viewBox="0 0 ${targetWidth} ${targetHeight}"
        shape-rendering="geometricPrecision"
        text-rendering="geometricPrecision"
      >
        ${svgElements.join("")}
      </svg>
    `);

      // Create a temporary sharp instance for the SVG
      const svgImage = sharp(svgBuffer, {
        density: Math.round(300 * upscaleFactor),
      }).resize(targetWidth, targetHeight, {
        fit: "fill",
        withoutEnlargement: true,
      });

      // Composite the SVG onto the upscaled image
      const composited = await sharp(upscaledImage)
        .composite([
          {
            input: await svgImage.toBuffer(),
            top: 0,
            left: 0,
          },
        ])
        .toBuffer();

      // Downscale back to original size with quality settings
      return sharp(composited)
        .resize(metadata.width, metadata.height, {
          kernel: "lanczos3",
          fit: "fill",
          withoutEnlargement: false,
        })
        .toBuffer();
    } catch (error) {
      console.error("❌ Error drawing text:", error);
      throw error;
    }
  }

  // Function to read Excel files with color support
  async function getAllTextFiles(dir = textDir) {
    try {
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir);
        console.log(
          "Created text directory. Please add Excel files and run again."
        );
        return [];
      }

      const textContents = [];
      const items = fs.readdirSync(dir);

      for (const item of items) {
        // Skip Excel temporary files (those starting with ~$)
        if (item.startsWith("~$")) {
          continue;
        }

        const fullPath = path.join(dir, item);
        const stat = fs.statSync(fullPath);

        if (stat.isDirectory()) {
          // Recursively get Excel files from subdirectories
          const subTextContents = await getAllTextFiles(fullPath);
          textContents.push(...subTextContents);
        } else {
          const ext = path.extname(item).toLowerCase();

          if (ext === ".xlsx" || ext === ".xls") {
            try {
              const workbook = new ExcelJS.Workbook();
              await workbook.xlsx.readFile(fullPath);

              // Process first worksheet only
              const worksheet = workbook.worksheets[0];

              // Create processColor function with workbook context
              const processColor = (font) => {
                if (!font || !font.color) return null;

                // Handle RGB color
                if (font.color.rgb) {
                  // Ensure RGB is 6 characters long
                  const rgb = font.color.rgb.padStart(6, "0");
                  return `#${rgb}`;
                }

                // Handle ARGB color
                if (font.color.argb) {
                  // Remove alpha channel and ensure remaining RGB is 6 characters
                  const rgb = font.color.argb.substring(2).padStart(6, "0");
                  return `#${rgb}`;
                }

                // Handle theme colors
                if (font.color.theme !== undefined) {
                  const themeColor = getThemeColor(workbook, font.color.theme);
                  if (font.color.tint !== undefined) {
                    return applyTint(themeColor, font.color.tint);
                  }
                  return themeColor;
                }

                return "#000000"; // Default color
              };

              // Read cells row by row
              worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                let rowText = "";
                let colorSegments = [];
                let boldSegments = [];
                let currentPosition = 0;

                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                  let cellValue = "";
                  let isBold = false;

                  if (cell.font && cell.font.bold) {
                    isBold = true;
                  }

                  switch (cell.type) {
                    case ExcelJS.ValueType.RichText:
                      cellValue = "";
                      cell.value.richText.forEach((rt) => {
                        const text = rt.text;
                        const startPos = currentPosition + cellValue.length;

                        if (rt.font && rt.font.bold) {
                          boldSegments.push({
                            start: startPos,
                            length: text.length,
                          });
                        }

                        if (rt.font && rt.font.color) {
                          const color = processColor(rt.font);
                          if (color) {
                            colorSegments.push({
                              start: startPos,
                              length: text.length,
                              color: color,
                            });
                          }
                        }

                        cellValue += text;
                      });
                      break;
                    default:
                      cellValue = cell.text || String(cell.value || "");

                      // Handle cell-level color
                      if (cell.font && cell.font.color) {
                        const color = processColor(cell.font);
                        if (color) {
                          colorSegments.push({
                            start: currentPosition,
                            length: cellValue.length,
                            color: color,
                          });
                        }
                      }
                  }

                  if (cellValue.trim()) {
                    if (rowText) {
                      rowText += " ";
                      currentPosition += 1;
                    }

                    if (isBold) {
                      boldSegments.push({
                        start: currentPosition,
                        length: cellValue.length,
                      });
                    }

                    rowText += cellValue;
                    currentPosition += cellValue.length;
                  }
                });

                if (rowText.trim()) {
                  textContents.push({
                    text: rowText.trim(),
                    colorSegments,
                    boldSegments,
                  });
                }
              });
            } catch (excelError) {
              console.warn(
                `⚠️ Skipped Excel file ${item}: ${excelError.message}`
              );
            }
          }
        }
      }

      return textContents;
    } catch (error) {
      console.error("❌ Error accessing directory:", error);
      return [];
    }
  }

  // Update the getThemeColor function
  function getThemeColor(workbook, theme) {
    // Updated default theme colors based on Office defaults
    const defaultThemeColors = [
      "#FFFFFF", // 0: Background1
      "#000000", // 1: Text1
      "#EEECE1", // 2: Background2
      "#1F497D", // 3: Text2
      "#4F81BD", // 4: Accent1
      "#C0504D", // 5: Accent2
      "#9BBB59", // 6: Accent3
      "#8064A2", // 7: Accent4
      "#4BACC6", // 8: Accent5
      "#F79646", // 9: Accent6
      "#0000FF", // 10: Hyperlink
      "#800080", // 11: Followed Hyperlink
    ];

    try {
      if (
        workbook.theme &&
        workbook.theme.themeElements &&
        workbook.theme.themeElements.clrScheme
      ) {
        const themeColors = workbook.theme.themeElements.clrScheme;

        // Map theme index to color scheme element
        const themeElements = [
          "lt1",
          "dk1",
          "lt2",
          "dk2",
          "accent1",
          "accent2",
          "accent3",
          "accent4",
          "accent5",
          "accent6",
          "hlink",
          "folHlink",
        ];

        const element = themeColors[themeElements[theme]];
        if (element) {
          if (element.rgb) {
            return `#${element.rgb}`;
          } else if (element.sysClr && element.sysClr.lastClr) {
            return `#${element.sysClr.lastClr}`;
          }
        }
      }
      return defaultThemeColors[theme] || "#000000";
    } catch (error) {
      console.error("Error getting theme color:", error);
      return "#000000";
    }
  }

  // Update the applyTint function for more accurate tinting
  function applyTint(hexColor, tint) {
    // Convert hex to RGB
    const r = parseInt(hexColor.slice(1, 3), 16);
    const g = parseInt(hexColor.slice(3, 5), 16);
    const b = parseInt(hexColor.slice(5, 7), 16);

    // Convert RGB to HSL
    const [h, s, l] = rgbToHsl(r, g, b);

    // Apply tint by adjusting luminance
    let newL;
    if (tint < 0) {
      newL = l * (1 + tint);
    } else {
      newL = l + (1 - l) * tint;
    }

    // Convert back to RGB
    const [newR, newG, newB] = hslToRgb(h, s, newL);

    // Convert to hex
    const toHex = (n) => {
      const hex = Math.round(Math.max(0, Math.min(255, n))).toString(16);
      return hex.length === 1 ? "0" + hex : hex;
    };

    return `#${toHex(newR)}${toHex(newG)}${toHex(newB)}`;
  }

  // Add helper functions for color space conversion
  function rgbToHsl(r, g, b) {
    r /= 255;
    g /= 255;
    b /= 255;

    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h,
      s,
      l = (max + min) / 2;

    if (max === min) {
      h = s = 0;
    } else {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

      switch (max) {
        case r:
          h = (g - b) / d + (g < b ? 6 : 0);
          break;
        case g:
          h = (b - r) / d + 2;
          break;
        case b:
          h = (r - g) / d + 4;
          break;
      }
      h /= 6;
    }

    return [h, s, l];
  }

  function hslToRgb(h, s, l) {
    let r, g, b;

    if (s === 0) {
      r = g = b = l;
    } else {
      const hue2rgb = (p, q, t) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
      };

      const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      const p = 2 * l - q;

      r = hue2rgb(p, q, h + 1 / 3);
      g = hue2rgb(p, q, h);
      b = hue2rgb(p, q, h - 1 / 3);
    }

    return [r * 255, g * 255, b * 255];
  }

  // Update processImage function to maintain source folder structure
  async function processImage(imagePath) {
    try {
      // Get relative path from images directory
      const relativePath = path.relative(
        imagesDir,
        path.dirname(path.join(imagesDir, imagePath))
      );
      const imageBuffer = readImageFile(imagePath);

      // Get text files from corresponding subfolder in text directory
      const textSubDir = path.join(textDir, relativePath);
      const textFiles = fs.existsSync(textSubDir)
        ? await getAllTextFiles(textSubDir)
        : await getAllTextFiles(textDir); // Fallback to root text dir if subfolder doesn't exist

      if (textFiles.length === 0) {
        console.log(`⚠️ No text files available for image: ${imagePath}`);
        return false;
      }

      const baseImageName = path.parse(imagePath).name;
      // Create result directory path maintaining subfolder structure
      const resultSubDir = path.join(resultDir, relativePath, baseImageName);

      // Create result subdirectory if it doesn't exist
      if (!fs.existsSync(resultSubDir)) {
        fs.mkdirSync(resultSubDir, { recursive: true });
      }

      // Get empty regions and metadata
      const [metadata, emptyRegions] = await Promise.all([
        sharp(imageBuffer).metadata(),
        findEmptyRegions(imageBuffer),
      ]);

      // Process each text content
      const processPromises = textFiles.map(async (textContent, i) => {
        const outputDir = path.join(resultSubDir, `text_${i + 1}`);

        // Create text-specific output directory if it doesn't exist
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir, { recursive: true });
        }

        try {
          const result = await insertTextIntoRegions(
            imageBuffer,
            textContent,
            emptyRegions
          );
          const outputPath = path.join(outputDir, `result.${OUTPUT_FORMAT}`);

          if (OUTPUT_FORMAT === "jpeg") {
            await sharp(result)
              .jpeg({
                quality: JPEG_QUALITY,
                mozjpeg: true,
                chromaSubsampling: "4:4:4",
              })
              .toFile(outputPath);
          } else {
            await sharp(result)
              .png({
                compressionLevel: PNG_COMPRESSION,
                palette: true,
              })
              .toFile(outputPath);
          }

          return true;
        } catch (error) {
          console.error(`❌ Failed processing text ${i + 1}: ${error.message}`);
          return false;
        }
      });

      const results = await Promise.all(processPromises);
      return results.some(Boolean);
    } catch (error) {
      console.error(`❌ Failed to process ${imagePath}: ${error.message}`);
      return false;
    }
  }

  // Update processBatch function to clean up logging
  async function processBatch(images, batchSize = 4) {
    const batches = [];
    for (let i = 0; i < images.length; i += batchSize) {
      batches.push(images.slice(i, i + batchSize));
    }

    let completed = 0;
    let successful = 0;

    console.log(
      `\n🚀 Processing ${images.length} images (batch size: ${batchSize})`
    );

    const progressBar = createProgressBar(images.length, "Progress");
    progressBar.update(0);

    for (const batch of batches) {
      const results = await Promise.all(
        batch.map((image) => processImage(image))
      );
      completed += batch.length;
      successful += results.filter(Boolean).length;

      progressBar.update(completed);
    }

    progressBar.complete();

    if (images.length > 0) {
      console.log(
        `\n✨ Complete: ${successful}/${images.length} successful (${(
          (successful / images.length) *
          100
        ).toFixed(1)}%)`
      );
    }
  }

  // Update processAllImages function to clean up logging
  async function processAllImages() {
    const images = getAllImages();
    if (images.length === 0) {
      console.log("⚠️ No images found in the images directory");
      return;
    }

    await processBatch(images);

    // Cleanup
    imageCache.clear();
    currentCacheSize = 0;
  }

  // Run the program using an immediately invoked async function
  (async () => {
    try {
      checkRequiredDirectories();
      await processAllImages();
    } catch (error) {
      console.error("❌ Error processing images:", error.message);
      process.exit(1);
    }
  })();

  // Update calculateTextWidth without BBCode handling
  function calculateTextWidth(text, fontSize) {
    const widthRatios = {
      default: 0.55,
      narrow: 0.35,
      wide: 0.75,
      space: 0.3,
    };

    const narrowChars = new Set("ijl,.'\"|!()[]{}/-_");
    const wideChars = new Set("mwWM@QOCDG%&#AHNUX");
    const spaceChars = new Set(" \t");

    let totalWidth = 0;
    let prevChar = "";

    for (const char of text) {
      if (spaceChars.has(char)) {
        totalWidth += fontSize * widthRatios.space;
      } else if (narrowChars.has(char)) {
        totalWidth += fontSize * widthRatios.narrow;
      } else if (wideChars.has(char)) {
        totalWidth += fontSize * widthRatios.wide;
      } else {
        totalWidth += fontSize * widthRatios.default;
      }

      prevChar = char;
    }

    return totalWidth;
  }

  // Update calculateOptimalFontSize function
  async function calculateOptimalFontSize(
    imageBuffer,
    region,
    text,
    paragraphs
  ) {
    try {
      // Calculate effective region dimensions (85% of width for margins)
      const effectiveWidth = region.width * 0.85;
      const effectiveHeight = region.height * 0.9;

      // Count total characters and bold characters
      let totalChars = 0;
      let boldChars = 0;
      paragraphs.forEach((para) => {
        totalChars += para.text.length;
        para.boldSegments.forEach((seg) => {
          boldChars += seg.length;
        });
      });

      // Calculate bold ratio
      const boldRatio = boldChars / totalChars;

      // Initial size estimation adjusted for bold ratio
      const initialSize = Math.min(
        Math.floor(effectiveHeight / (paragraphs.length * 1.5)),
        Math.floor(effectiveWidth / (20 + boldRatio * 5)) // Adjust width estimation based on bold ratio
      );

      // Binary search parameters
      let minSize = 14;
      let maxSize = 72;
      let bestSize = initialSize;
      let bestLines = [];
      let iterations = 0;
      const maxIterations = 10;

      while (minSize <= maxSize && iterations < maxIterations) {
        const fontSize = Math.floor((minSize + maxSize) / 2);
        const boldFontSize = fontSize * 1.1; // Bold text is 10% larger

        // Test wrap with current font sizes
        const testLines = [];
        let totalHeight = 0;
        const lineHeight = fontSize * 1.2;

        // Simulate text wrapping with mixed font sizes
        let fitsFailed = false;

        paragraphs.forEach((para, index) => {
          if (index > 0) {
            totalHeight += lineHeight * 0.5;
          }

          let currentLine = "";
          let lineWidth = 0;
          let position = 0;

          // Process each character with its appropriate font size
          for (const char of para.text) {
            const isBold = para.boldSegments.some(
              (seg) =>
                position >= seg.start && position < seg.start + seg.length
            );
            const charWidth = calculateTextWidth(
              char,
              isBold ? boldFontSize : fontSize
            );

            if (lineWidth + charWidth <= effectiveWidth) {
              currentLine += char;
              lineWidth += charWidth;
            } else {
              if (currentLine) {
                testLines.push(currentLine);
                totalHeight += lineHeight;
                if (totalHeight > effectiveHeight) {
                  fitsFailed = true;
                  break;
                }
              }
              currentLine = char;
              lineWidth = charWidth;
            }
            position++;
          }

          if (fitsFailed) return;

          if (currentLine) {
            testLines.push(currentLine);
            totalHeight += lineHeight;
            if (totalHeight > effectiveHeight) {
              fitsFailed = true;
            }
          }
        });

        // Check if this size fits well
        if (
          !fitsFailed &&
          totalHeight <= effectiveHeight &&
          testLines.length > 0
        ) {
          // This size fits, try larger
          bestSize = fontSize;
          bestLines = testLines;
          minSize = fontSize + 1;
        } else {
          // Too big, try smaller
          maxSize = fontSize - 1;
        }

        iterations++;
      }

      // Calculate final metrics
      const finalLineHeight = bestSize * 1.2;
      const totalTextHeight = bestLines.length * finalLineHeight;
      const fillRatio = totalTextHeight / effectiveHeight;

      return {
        normalFontSize: bestSize,
        boldFontSize: bestSize * 1.1,
        lineSpacingRatio: 1.2,
        estimatedLines: bestLines.length,
      };
    } catch (error) {
      console.error("❌ Error calculating font size:", error.message);
      return {
        normalFontSize: 20,
        boldFontSize: 22,
        lineSpacingRatio: 1.2,
        estimatedLines: paragraphs.length,
      };
    }
  }

  // After wrapping - Optimize layout based on actual content
  function optimizeLayout(wrappedLines, region, lineHeight, fontSize) {
    // Calculate actual content metrics
    const contentHeight = wrappedLines.reduce((height, line) => {
      if (line.isParagraphSpace) {
        return height + lineHeight * 0.5; // Paragraph spacing
      }
      return height + lineHeight;
    }, 0);

    // Calculate content density
    const contentDensity =
      wrappedLines.reduce((sum, line) => {
        if (!line.text || line.isParagraphSpace) return sum;
        return sum + (line.text.length * fontSize) / region.width;
      }, 0) / wrappedLines.length;

    let adjustedLineHeight = lineHeight;
    let adjustedFontSize = fontSize;

    // Adjust based on content height ratio
    const heightRatio = contentHeight / region.height;

    if (heightRatio > 0.95) {
      // Content is too tall
      if (heightRatio > 1.1) {
        // Significantly too tall - reduce both spacing and font
        adjustedLineHeight = lineHeight * 0.9;
        adjustedFontSize = fontSize * 0.95;
      } else {
        // Slightly too tall - reduce spacing only
        adjustedLineHeight = lineHeight * 0.95;
      }
    } else if (heightRatio < 0.7 && contentDensity < 0.5) {
      // Content is too short and sparse - increase spacing
      adjustedLineHeight = lineHeight * 1.1;
    }

    // Calculate vertical positioning
    const finalContentHeight = wrappedLines.reduce((height, line) => {
      if (line.isParagraphSpace) {
        return height + adjustedLineHeight * 0.5;
      }
      return height + adjustedLineHeight;
    }, 0);

    const startY =
      region.y +
      Math.max(
        adjustedLineHeight * 0.2,
        (region.height - finalContentHeight) / 2
      );

    return {
      lineHeight: adjustedLineHeight,
      fontSize: adjustedFontSize,
      startY: startY,
      contentDensity: contentDensity,
    };
  }

  // Remove the complex line spacing calculation
  function calculateLineSpacingRatio(fontSize, lineCount) {
    return 1.2; // Fixed line spacing ratio
  }

  // Add this helper function to determine optimal vertical alignment
  function determineVerticalAlignment(region, totalTextHeight, textLines) {
    // Calculate content density (text characters per line average)
    const contentDensity =
      textLines.reduce(
        (sum, line) => sum + (line.text ? line.text.length : 0),
        0
      ) / textLines.length;

    // Calculate height ratio
    const heightRatio = totalTextHeight / region.height;

    // Determine if content is sparse
    const isSparseContent = contentDensity < 30 && textLines.length < 5;

    // If region is much larger than text, center it
    if (heightRatio < 0.5 || isSparseContent) {
      return "center";
    }

    // If text nearly fills the region, use top alignment
    if (heightRatio > 0.8) {
      return "top";
    }

    // For medium-length text, check content characteristics
    const hasLongLines = textLines.some(
      (line) => line.text && line.text.length > 50
    );
    return hasLongLines ? "top" : "center";
  }

  // Update insertTextIntoRegions function to remove unnecessary logging
  async function insertTextIntoRegions(imageBuffer, textContent, emptyRegions) {
    try {
      const text = textContent.text || String(textContent);

      if (!text) {
        console.log("⚠️ No text content to insert");
        return imageBuffer;
      }

      const regions = emptyRegions.filter(
        (r) => r.width >= 100 && r.height >= 30
      );
      if (regions.length === 0) {
        console.log("⚠️ No valid regions found for text insertion");
        return imageBuffer;
      }

      const region = regions[0];

      // Process text with formatting
      const paragraphs = text.split("\n").map((para, index) => {
        const paraStart = text.indexOf(para);

        // Adjust segments for this paragraph
        const paraColorSegments = textContent.colorSegments
          .filter(
            (seg) =>
              seg.start >= paraStart && seg.start < paraStart + para.length
          )
          .map((seg) => ({
            ...seg,
            start: seg.start - paraStart,
          }));

        const paraBoldSegments = textContent.boldSegments
          .filter(
            (seg) =>
              seg.start >= paraStart && seg.start < paraStart + para.length
          )
          .map((seg) => ({
            ...seg,
            start: seg.start - paraStart,
          }));

        return {
          text: para,
          colorSegments: paraColorSegments,
          boldSegments: paraBoldSegments,
        };
      });

      // Calculate font sizes with the full text content
      const { normalFontSize, boldFontSize, estimatedLines } =
        await calculateOptimalFontSize(imageBuffer, region, text, paragraphs);

      // Calculate spacing metrics with 80% width
      const effectiveWidth = region.width * 0.85;
      const textStartX = region.x + region.width * 0.075;
      const lineHeight = normalFontSize * 1.2;

      // Process each paragraph and wrap text
      const wrappedLines = [];
      let isFirstParagraph = true;

      paragraphs.forEach((paragraph, pIndex) => {
        // Add paragraph spacing except for first paragraph
        if (!isFirstParagraph) {
          wrappedLines.push({
            text: "",
            color: "#000000",
            isNewParagraph: true,
            isBold: false,
            isParagraphSpace: true,
          });
        }
        isFirstParagraph = false;

        // Split paragraph into words
        const words = paragraph.text.split(" ");
        let currentLine = "";
        let currentWidth = 0;
        let currentLineStart = paragraph.text.indexOf(words[0]);
        let currentLineColors = [];

        words.forEach((word, wIndex) => {
          const wordStart = paragraph.text.indexOf(word, currentLineStart);
          const isBold = paragraph.boldSegments.some(
            (seg) =>
              seg.start <= wordStart &&
              seg.start + seg.length >= wordStart + word.length
          );

          // Get color for the word
          const wordColor =
            paragraph.colorSegments.find(
              (seg) =>
                seg.start <= wordStart &&
                seg.start + seg.length >= wordStart + word.length
            )?.color || "#000000";

          const wordWidth = calculateTextWidth(
            word,
            isBold ? boldFontSize : normalFontSize
          );
          const spaceWidth = calculateTextWidth(" ", normalFontSize);

          if (currentWidth + wordWidth <= effectiveWidth) {
            // Add word to current line with its color
            if (currentLine) {
              currentLine += " ";
              currentWidth += spaceWidth;
            }
            currentLine += word;
            currentWidth += wordWidth;

            // Store color information for the word
            currentLineColors.push({
              text: word,
              start: currentLine.length - word.length,
              color: wordColor,
            });
          } else {
            // Start new line
            if (currentLine) {
              wrappedLines.push({
                text: currentLine,
                isNewParagraph: false,
                isBold: paragraph.boldSegments.some(
                  (seg) =>
                    seg.start <= currentLineStart &&
                    seg.start + seg.length >=
                      currentLineStart + currentLine.length
                ),
                colorSegments: currentLineColors,
              });
            }
            currentLine = word;
            currentWidth = wordWidth;
            currentLineStart = wordStart;
            currentLineColors = [
              {
                text: word,
                start: 0,
                color: wordColor,
              },
            ];
          }
        });

        // Add last line of paragraph with color segments
        if (currentLine) {
          wrappedLines.push({
            text: currentLine,
            isNewParagraph: false,
            isBold: paragraph.boldSegments.some(
              (seg) =>
                seg.start <= currentLineStart &&
                seg.start + seg.length >= currentLineStart + currentLine.length
            ),
            colorSegments: currentLineColors,
          });
        }

        // Remove the paragraph separator at the end of paragraphs
        if (pIndex < paragraphs.length - 1) {
          // We don't need this anymore as we're handling spacing differently
          // wrappedLines.push({
          //   text: '',
          //   color: '#000000',
          //   isNewParagraph: true,
          //   isBold: false
          // });
        }
      });

      // Calculate paragraph spacing after wrappedLines is created
      const paragraphSpacing =
        lineHeight * (wrappedLines.length > 10 ? 1.1 : 1.3);

      // Calculate total text height
      const totalTextHeight = wrappedLines.reduce((height, line) => {
        if (line.isParagraphSpace) {
          return (
            height +
            (wrappedLines.length > 10 ? lineHeight * 0.5 : paragraphSpacing)
          );
        }
        return height + lineHeight;
      }, 0);

      // Determine vertical alignment
      const verticalAlignment = determineVerticalAlignment(
        region,
        totalTextHeight,
        wrappedLines
      );

      // Calculate starting Y position based on alignment
      let startY;
      const topPadding = lineHeight * 0.5;
      const bottomPadding = lineHeight * 0.5;

      switch (verticalAlignment) {
        case "top":
          startY = region.y + topPadding;
          break;
        case "bottom":
          startY = region.y + region.height - totalTextHeight - bottomPadding;
          break;
        case "center":
        default:
          startY = region.y + (region.height - totalTextHeight) / 2;
          break;
      }

      // Ensure startY is within bounds
      startY = Math.max(
        region.y + topPadding,
        Math.min(
          startY,
          region.y + region.height - totalTextHeight - bottomPadding
        )
      );

      // Create text objects with updated positioning
      let currentY = startY;
      const textObjects = wrappedLines
        .map((line) => {
          if (!line.text && !line.isNewParagraph) return null;

          if (line.isParagraphSpace) {
            currentY += paragraphSpacing - lineHeight;
            return null;
          }

          const fontSize = line.isBold ? boldFontSize : normalFontSize;
          const textObject = {
            type: "text",
            text: line.text,
            x: textStartX + effectiveWidth / 2,
            y: currentY + fontSize / 2,
            color: line.color || "#000000",
            fontSize: fontSize,
            textAnchor: "middle",
            dominantBaseline: "middle",
            fontWeight: line.isBold ? "bold" : "normal",
            fontFamily: "Arial, Helvetica, sans-serif",
            style: `letter-spacing: 0px`,
            colorSegments: line.colorSegments || [], // Add color segments information
          };

          currentY += lineHeight;
          return textObject;
        })
        .filter((obj) => obj !== null);

      // Draw the text
      return await drawRegionsWithText(imageBuffer, [region], textObjects);
    } catch (error) {
      console.error("❌ Text insertion failed:", error.message);
      throw error;
    }
  }

  // Add these constants at the top of the file
  const JPEG_QUALITY = 90; // Adjust JPEG quality (0-100)
  const PNG_COMPRESSION = 9; // PNG compression level (0-9)
  const OUTPUT_FORMAT = "jpeg"; // or 'png' based on your needs

  // Add these utility functions for the progress bar
  function createProgressBar(total, title = "Progress") {
    const barWidth = 30;
    let current = 0;

    function update(value) {
      current = value;
      const percentage = (current / total) * 100;
      const filled = Math.round((barWidth * current) / total);
      const empty = barWidth - filled;

      const filledBar = "█".repeat(filled);
      const emptyBar = "░".repeat(empty);
      const percentageStr = percentage.toFixed(1).padStart(5);

      readline.clearLine(process.stdout, 0);
      readline.cursorTo(process.stdout, 0);
      process.stdout.write(
        `${consoleStyles.info}${title}: ${consoleStyles.reset}` +
          `[${consoleStyles.success}${filledBar}${consoleStyles.reset}${emptyBar}] ` +
          `${percentageStr}% (${current}/${total})`
      );
    }

    function complete() {
      readline.clearLine(process.stdout, 0);
      readline.cursorTo(process.stdout, 0);
      process.stdout.write(
        `${consoleStyles.info}${title}: ${consoleStyles.reset}` +
          `[${consoleStyles.success}${"█".repeat(barWidth)}${
            consoleStyles.reset
          }] ` +
          `100.0% (${total}/${total})\n`
      );
    }

    return { update, complete };
  }
};

module.exports = { run };
