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

    function test4() {
        test('https://raw.githubusercontent.com/MiroHuang/Miro/master/TableWithDirectFormat.xml');
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

    function insertImage() {
        var imgHTML = "<img " + "src='https://www.baidu.com/img/bd_logo1.png' " + "alt='上海鲜花港 - 郁金香' img />";
        Office.context.document.setSelectedDataAsync(
            imgHTML, { coercionType: Office.CoercionType.Html },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    write('Error: ' + asyncResult.error.message);
                }
            });
    }

    function insertImage1() {
        var imgHTML = "<img " + "src='https://localhost:44342/OOXMLSamples/ChartMarkup.xml' " + "alt='上海鲜花港 - 郁金香' img />";
        Office.context.document.setSelectedDataAsync(
            imgHTML, { coercionType: Office.CoercionType.Html },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    write('Error: ' + asyncResult.error.message);
                }
            });
    }
    /**
     * get the ooxml of the doc
     * 
     * 
     * @memberOf WordDocumentService
     */
    function getOoxml() {
        // Run a batch operation against the Word object model.
        return Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to get the HTML contents of the body.
            var bodyOOXML = body.getOoxml();

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            // return context.sync().then(function () {
            //     console.log("Body HTML contents: " + bodyHTML.value);
            //     return bodyHTML.value;
            // });
            //return context.sync().then(() => { return bodyOOXML.value });
            return context.sync().then(function () {
                console.log("Body HTML contents: " + bodyHTML.value);
                return bodyHTML.value;
            });
        })
            .catch(function (error) {
                console.log("Error: " + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
                return "";
            });
    }

    /**
     * set the ooxml of the doc
     * 
     * @param {string} ooxml 
     * 
     * @memberOf WordDocumentService
     */
    function setOoxml(ooxml) {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to insert OOXML in to the beginning of the body.
            body.insertOoxml(ooxml, Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('OOXML added to the beginning of the document body.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function applyTemplate(template) {
        this.wordDocument.setOoxml(template.TemplateContent);
    }

    //使用OOXML方式插入图像
    function setOOXMLImage(imgOOXML) {
        if (Office.CoercionType.Ooxml) {
            Office.context.document.setSelectedDataAsync(
                imgOOXML,
                { coercionType: "ooxml" },
                function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        write(asyncResult.error.message);
                    }
                });
        }
        else {
            write("Setting data as Open XML is not supported ");
            //writeError("Setting data as Open XML is not supported ");
        }
    }

    function getText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            { valueFormat: "unformatted", filterType: "all" },
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    write(error.name + ": " + error.message);
                }
                else {
                    // Get selected data.
                    var dataValue = asyncResult.value;
                    write('Selected data is ' + dataValue);
                }
            });
    }

    // Function that writes to a div with id='message' on the page.
    function write(message) {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the start of the document body.
            body.insertText(message, Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Anton Chekhov.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function insertEmersonQuoteAtSelection() {
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = thisDocument.getSelection();

            // Queue a command to replace the selected text.
            range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Ralph Waldo Emerson.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function insertChekhovQuoteAtTheBeginning() {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the start of the document body.
            body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Anton Chekhov.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function insertChineseProverbAtTheEnd() {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the end of the document body.
            body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from a Chinese proverb.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }
})();