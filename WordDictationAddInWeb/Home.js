/// <reference path="/Scripts/FabricUI/MessageBanner.js" />
/// <reference path="distrib/speech.browser.sdk.js" />


var SpeechSDK = null;
var _recognizer = null;
var _context = null;

(function () {
    "use strict";

    var messageBanner;

    function InitalizeSdk(OnComplete)
    {
    	require(["Speech.Browser.Sdk"], function (SDK) {
    		OnComplete(SDK);
    	}, function (error) {
    		console.error(error);
    	});
    }

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
    	$(document).ready(function () {
    		// Initialize the FabricUI notification mechanism and hide it
    		var element = document.querySelector('.ms-MessageBanner');
    		messageBanner = new fabric.MessageBanner(element);
    		messageBanner.hideBanner();

    		$("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
    		$('#button-text').text("Start Dictation");
    		$('#button-stop-text').text("Stop Dictation");
    		$('#button-desc').text("Highlights the longest word.");

    		// Add a click event handler for the highlight button.
    		$('#highlight-button').click(
				StartDictation);

    		$('#stop-button').click(
			StopDictation);

    		$('#highlight-button').disabled = true;

    		InitalizeSdk(function (SDK) {
    			SpeechSDK = SDK;
    			$('#highlight-button').disabled = false;
    		})

    	});
    };

    function UpdateStatus(status) {
    	var div = document.getElementById("statusDiv");
    	div.innerText = status;
    }

    function hightlightLongestWord() {

        Word.run(function (context) {

            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // variable for keeping the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = context.document.body.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');

                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync)
        })
        .catch(errorHandler);
    } 

    function RecognizerSetup(SDK) {


    	var recognitionMode = SDK.RecognitionMode.Dictation;
    	var language = "en-US";
    	var format = SDK.SpeechResultFormat.Simple;
    	var subscriptionKey = "70aa2103f3664071800b51ffcfdcd3ad";

    	var recognizerConfig = new SDK.RecognizerConfig(
			new SDK.SpeechConfig(
				new SDK.Context(
					new SDK.OS(navigator.userAgent, "Browser", null),
					new SDK.Device("SpeechSample", "SpeechSample", "1.0.00000"))),
			recognitionMode,
			language, // Supported laguages are specific to each recognition mode. Refer to docs.
			format); // SDK.SpeechResultFormat.Simple (Options - Simple/Detailed)

    	// Alternatively use SDK.CognitiveTokenAuthentication(fetchCallback, fetchOnExpiryCallback) for token auth
    	var authentication = new SDK.CognitiveSubscriptionKeyAuthentication(subscriptionKey);

  
    	return SDK.CreateRecognizer(recognizerConfig, authentication);
    	
    }

    function UpdateText(text, fFinal)
    {
    	Word.run(function (context) {

    		var contentControl = null;
    		var updateID = false;
    		if (fFinal)
    		{
    			if(_contentControlID != null)
    			{
    				contentControl = context.document.contentControls.getById(_contentControlID);
    				contentControl.insertText(text + " ", "Replace");
    				var range = contentControl.getRange("End");
    				contentControl.delete(true);
    				_contentControlID = null;
    				range.select("End");

    				return context.sync().then(function () {
    					console.log('Text added to the end of the range.');
    				});
    			}
    			else
    			{
    				var range = context.document.getSelection();
    				range.insertText(text + " ", "Replace");
    				range.select("End");

    				return context.sync().then(function () {
    					console.log('Text added to the end of the range.');
    				});
    			}
    		}
    		else
    		{
    			if (_contentControlID == null) {
    				var range = context.document.getSelection();
    				contentControl = range.insertContentControl();
    				contentControl.appearance = "hidden";


    				context.load(contentControl, "id");
    			}
    			else {
    				contentControl = context.document.contentControls.getById(_contentControlID);
    			}

    			contentControl.insertText(text + " ", "End");

    			return context.sync().then(function () {
    				if(_contentControlID == null)
    				{
    					_contentControlID = contentControl.id;
    				}
    			});
    		}
    	});
    	
    }

	// Start the recognition
    function RecognizerStart(SDK, recognizer) {
    	recognizer.Recognize((event) => {
    		/*
			 Alternative syntax for typescript devs.
			 if (event instanceof SDK.RecognitionTriggeredEvent)
			*/

    		UpdateStatus(JSON.stringify(event));

    		switch (event.Name) {
    			case "RecognitionTriggeredEvent":
    				break;
    			case "ListeningStartedEvent":
    				break;
    			case "RecognitionStartedEvent":
    				break;
    			case "SpeechStartDetectedEvent":
    				break;
    			case "SpeechHypothesisEvent":
    				break;
    			case "SpeechFragmentEvent":
    				UpdateText(event.Result.Text, false);
    				break;
    			case "SpeechEndDetectedEvent":
    				break;
    			case "SpeechSimplePhraseEvent":
    				if (event.Result.RecognitionStatus == "Success") {
    					UpdateText(event.Result.DisplayText, true);
    				}
    				break;
    			case "SpeechDetailedPhraseEvent":
    				break;
    			case "RecognitionEndedEvent":
    				break;
    			default:
    				
    				break;
    		}
    	})
		.On(() => {
			// The request succeeded. Nothing to do here.
		},
		(error) => {
			console.error(error);
		});
    }

	// Stop the Recognition.
    function RecognizerStop(SDK, recognizer) {
    	// recognizer.AudioSource.Detach(audioNodeId) can be also used here. (audioNodeId is part of ListeningStartedEvent)
    	recognizer.AudioSource.TurnOff();
    }

    var _contentControlID = null;
    function StartDictationCore() {

    	Word.run(function (context) {
    		if (_recognizer == null) {
    			_recognizer = RecognizerSetup(SpeechSDK);
    		}
    		else {
    			RecognizerStop(SpeechSDK, _recognizer);
    		}

    		var range = context.document.getSelection();

    		RecognizerStart(SpeechSDK, _recognizer);
    		
    	});
    }

    function StartDictation()
    {
    	if (SpeechSDK != null)
    		StartDictationCore();
    	else
    	{
    		InitalizeSdk(function (SDK) {
    			SpeechSDK = SDK;
    			StartDictationCore();
    		})
    	}
    }

    function StopDictation()
    {
    	if (SpeechSDK != null && _recognizer != null)
    		RecognizerStop(SpeechSDK, _recognizer);
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
