import sharp from "sharp";

const resizeImage = async (imagePath, maxWidth, maxHeight) => {
  try {
    // Load the image
    const image = sharp(imagePath);

    // Get the metadata of the image to preserve aspect ratio
    const { width, height } = await image.metadata();

    // Calculate the new dimensions while preserving the aspect ratio
    let newWidth = width;
    let newHeight = height;

    if (width > height) {
      if (width > maxWidth) {
        newWidth = maxWidth;
        newHeight = Math.round((height * maxWidth) / width);
      }
    } else {
      if (height > maxHeight) {
        newHeight = maxHeight;
        newWidth = Math.round((width * maxHeight) / height);
      }
    }

    // Resize the image
    const resizedImage = await image.resize(newWidth, newHeight).toBuffer();

    // Create a new image for debugging purposes
    const newImage = sharp(resizedImage);

    // Save the resized image to disk
    await newImage.toFile("./xcv.png");

    // Optionally, you can save the resized image to a new file for debugging
    // await newImage.toFile('resized_image.jpg');

    return newImage; // Return the resized image
  } catch (error) {
    console.error("Error resizing image:", error);
    throw error; // Rethrow the error for handling upstream
  }
};

// Example usage
(async () => {
  const resized = await resizeImage("./big.jpg", 80, 60);
  // Further processing or saving of `resized` can be done here
})();
