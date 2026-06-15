// ...existing code...
const productImages = document.querySelectorAll('.product-image');
productImages.forEach(image => {
    image.style.maxWidth = '300px'; // Ensure consistent max width
    image.style.height = 'auto'; // Maintain aspect ratio
    image.style.objectFit = 'contain'; // Fit image within bounds
});
// ...existing code...
