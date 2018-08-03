'use strict';
'use Microsoft.Office.Core';

(function () {
    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                // Do something that is only available via the new APIs
                $('#emerson').click(test1);
                $('#checkhov').click(test2);
                $('#proverb').click(test3);
                $('#test').click(test4);
    
                $('#supportedVersion').html('This code is using Word 2016 or greater.');
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or greater.');
            }
        });
    };

    function test1() {
        test('https://raw.githubusercontent.com/MiroHuang/Miro/master/Chart.xml');
    }

    function test2() {
        test('https://raw.githubusercontent.com/MiroHuang/Miro/master/FormatForMarkup.xml');
    }

    function test3() {
        test('https://raw.githubusercontent.com/MiroHuang/Miro/master/SimpleImage.xml');
    }

    function test(filename) {
        var myOOXMLRequest = new XMLHttpRequest();
        var myXML;
        myOOXMLRequest.open('GET', filename, false);
        myOOXMLRequest.send();
        if (myOOXMLRequest.status === 200) {
            myXML = myOOXMLRequest.responseText;
        }
        Office.context.document.setSelectedDataAsync(
            myXML, { coercionType: 'ooxml' },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    write('Error: ' + asyncResult.error.message);
                }
            });
    }

    function test4() {
        var imgHTML = "<img " + "src='https://www.baidu.com/img/bd_logo1.png' " + "alt='上海鲜花港 - 郁金香' img />";
        Office.context.document.setSelectedDataAsync(
            imgHTML, { coercionType: Office.CoercionType.Html },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    write('Error: ' + asyncResult.error.message);
                }
            });
    }
})();