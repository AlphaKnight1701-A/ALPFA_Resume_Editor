document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (file) {
    console.log('File selected:', file.name);
    // You can add additional file processing logic here
  }
});