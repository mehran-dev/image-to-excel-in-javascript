import sharp from "sharp";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url); // get the resolved path to the file
const __dirname = path.dirname(__filename); // get the name of the directory
async function processImage(imagePath) {
  try {
    // Load the image using sharp and get metadata
    const image = sharp(imagePath);
    const { width, height, channels } = await image.metadata();

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Pixel Art");

    // Get the raw pixel data from the image (RGBA format)
    const rawData = await image.raw().toBuffer();

    // Iterate through each pixel and color the cell background
    for (let y = 0; y < height; y++) {
      const row = worksheet.getRow(y + 1); // Get the row for the current y-coordinate
      for (let x = 0; x < width; x++) {
        const idx = (y * width + x) * channels; // Calculate index for the pixel (4 values per pixel: RGBA)

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
    for (let i = 1; i <= width; i++) {
      worksheet.getColumn(i).width = 2; // Set column width to 2 for better visual representation
    }
    for (let i = 1; i <= height; i++) {
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
console.log(__dirname);

processImage(imagePath);
