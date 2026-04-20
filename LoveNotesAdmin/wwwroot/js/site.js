// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
window.downloadFile = (fileName, base64String) => {
    const link = document.createElement('a');
    link.download = fileName;
    link.href = "data:text/plain;base64," + base64String;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}