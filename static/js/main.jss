// static/js/main.js
document.addEventListener('DOMContentLoaded', function() {
  // Add any client-side functionality here
  
  // Example: Highlight the selected company folder
  const companyFolders = document.querySelectorAll('.company-folder');
  companyFolders.forEach(folder => {
    folder.addEventListener('click', function() {
      companyFolders.forEach(f => f.classList.remove('active'));
      this.classList.add('active');
    });
  });
});