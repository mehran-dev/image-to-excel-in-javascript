import sharp from "sharp";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url); // get the resolved path to the file
const __dirname = path.dirname(__filename); // get the name of the directory
// Function to save the image to disk
async function saveImageToDisk(imageBuffer, outputPath) {
  return new Promise((resolve, reject) => {
    fs.writeFile(outputPath, imageBuffer, (err) => {
      if (err) {
        return reject(err);
      }
      console.log(`Image saved at ${outputPath}`);
      resolve();
    });
  });
}

// Function to resize image while maintaining the aspect ratio
async function resizeImage(image, maxWidth, maxHeight) {
  const { width, height } = await image.metadata();

  // Check if the image needs resizing
  if (width > maxWidth || height > maxHeight) {
    console.log(`Image is too large (${width}x${height}). Resizing...`);

    // Calculate the scale factor while preserving aspect ratio
    const widthRatio = maxWidth / width;
    const heightRatio = maxHeight / height;
    const resizeRatio = Math.min(widthRatio, heightRatio); // Preserve aspect ratio

    // Calculate the new dimensions
    const newWidth = Math.round(width * resizeRatio);
    const newHeight = Math.round(height * resizeRatio);

    // Resize the image
    const resizedImage = image.resize(newWidth, newHeight);

    // Get the image buffer of the resized image
    const imageBuffer = await resizedImage.toBuffer();

    return { imageBuffer, newWidth, newHeight }; // Return the buffer and new dimensions
  }

  // Return the original image if no resizing is needed
  const imageBuffer = await image.toBuffer();
  return { imageBuffer, width, height }; // Return original image buffer and dimensions
}

async function processImage(imagePath) {
  try {
    // Load the image using sharp
    let image = sharp(imagePath);

    // Define max dimensions
    const maxWidth = 100;
    const maxHeight = 100;

    // Resize image while maintaining aspect ratio
    const { imageBuffer, newWidth, newHeight } = await resizeImage(
      image,
      maxWidth,
      maxHeight
    );

    // Define output path for the resized image
    const outputResizedPath = path.join(
      __dirname,
      "./image-pixel/resized-image.png"
    );

    // Save the resized image to disk
    await saveImageToDisk(imageBuffer, outputResizedPath);

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Pixel Art");

    // Get the raw pixel data from the resized image (RGBA format)
    const rawData = await sharp(imageBuffer).raw().toBuffer(); // Get raw data from buffer

    // Iterate through each pixel and color the cell background
    for (let y = 0; y < newHeight; y++) {
      const row = worksheet.getRow(y + 1); // Get the row for the current y-coordinate
      for (let x = 0; x < newWidth; x++) {
        const idx = (y * newWidth + x) * 4; // Calculate index for the pixel (4 values per pixel: RGBA)

        // Get RGBA values
        const r = rawData[idx]; // Red channel
        const g = rawData[idx + 1]; // Green channel
        const b = rawData[idx + 2]; // Blue channel
        const a = rawData[idx + 3]; // Alpha channel

        // Convert RGB to Hex format (ignore alpha for solid color)
        const hexColor = `#${((1 << 24) + (r << 16) + (g << 8) + b)
          .toString(16)
          .slice(1)}`;

        // Get the cell from the row and color the background
        const cell = row.getCell(x + 1);
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: hexColor.replace("#", "") }, // Set the color (fully opaque)
        };
      }
      row.commit(); // Commit the row after updating it
    }

    // Adjust the row height and column width to better represent the image
    for (let i = 1; i <= newWidth; i++) {
      worksheet.getColumn(i).width = 2; // Set column width to 2 for better visual representation
    }
    for (let i = 1; i <= newHeight; i++) {
      worksheet.getRow(i).height = 15; // Set row height
    }

    // Save the workbook
    const outputFilePath = path.join(
      __dirname,
      "./image-pixel/ImageAsExcel.xlsx"
    );
    await workbook.xlsx.writeFile(outputFilePath);
    console.log(`Excel file created at ${outputFilePath}`);
  } catch (error) {
    console.error("Error processing image:", error);
  }
}

// Example usage
const imagePath = "./image.png"; // Replace with your image path
processImage(imagePath);
