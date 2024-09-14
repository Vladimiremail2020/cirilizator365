function transliterate(text, toCyrillic) {
    const latin = ['A', 'a', 'B', 'b', 'C', 'c', 'Č', 'č', 'Ć', 'ć', 'D', 'd', 'Dž', 'dž', 'Đ', 'đ', 'E', 'e', 'F', 'f', 'G', 'g', 'H', 'h', 'I', 'i', 'J', 'j', 'K', 'k', 'L', 'l', 'Lj', 'lj', 'M', 'm', 'N', 'n', 'Nj', 'nj', 'O', 'o', 'P', 'p', 'R', 'r', 'S', 's', 'Š', 'š', 'T', 't', 'U', 'u', 'V', 'v', 'Z', 'z', 'Ž', 'ž'
]; // Add full alphabet
    const cyrillic = ['А', 'а', 'Б', 'б', 'Ц', 'ц', 'Ч', 'ч', 'Ћ', 'ћ', 'Д', 'д', 'Џ', 'џ', 'Ђ', 'ђ', 'Е', 'е', 'Ф', 'ф', 'Г', 'г', 'Х', 'х', 'И', 'и', 'Ј', 'ј', 'К', 'к', 'Л', 'л', 'Љ', 'љ', 'М', 'м', 'Н', 'н', 'Њ', 'њ', 'О', 'о', 'П', 'п', 'Р', 'р', 'С', 'с', 'Ш', 'ш', 'Т', 'т', 'У', 'у', 'В', 'в', 'З', 'з', 'Ж', 'ж'
]; // Add corresponding Cyrillic alphabet

    let from, to;
    if (toCyrillic) {
        from = latin;
        to = cyrillic;
    } else {
        from = cyrillic;
        to = latin;
    }

    let result = '';
    for (let char of text) {
        let index = from.indexOf(char);
        result += index !== -1 ? to[index] : char;
    }
    return result;
}

document.getElementById('transliterateButton').addEventListener('click', () => {
    Word.run(function (context) {
        var range = context.document.getSelection();
        range.load('text');
        return context.sync().then(function () {
            var transliteratedText = transliterate(range.text, true); // or false for Latin
            range.insertText(transliteratedText, 'Replace');
            return context.sync();
        });
    }).catch(function (error) {
        console.log('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
});
