const btn = document.querySelector('.talk');
const content = document.querySelector('.content');

btn.addEventListener('click', () => {
    content.textContent = "Listening...";

    fetch('/listen')
        .then(response => response.json())
        .then(data => {
            content.textContent = data.message;
        });
});
