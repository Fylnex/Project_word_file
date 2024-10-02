function generateDocument() {
    // Получение данных из формы
    let title = document.getElementById('title').value;
    let authors = document.getElementById('authors').value;
    let abstract = document.getElementById('abstract').value;
    let keywords = document.getElementById('keywords').value.split(',');

    let paragraph = document.getElementById('paragraph').value;
    let heading = document.getElementById('heading').value;
    let subheading = document.getElementById('subheading').value;

    let imageInput = document.getElementById('image');
    let imageCaption = document.getElementById('imageCaption').value;

    let codeContent = document.getElementById('codeContent').value;
    let codeTitle = document.getElementById('codeTitle').value;

    let numberedList = document.getElementById('numberedList').value.split('\n');
    let unorderedList = document.getElementById('unorderedList').value.split('\n');

    let references = document.getElementById('references').value.split('\n');

    // Отправка данных в Python
    eel.generate_document(
        title, authors, abstract, keywords,
        paragraph, heading, subheading,
        imageCaption, codeTitle, codeContent,
        numberedList, unorderedList, references
    )(function(response) {
        document.getElementById('message').innerText = response;
    });
}
