/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, OfficeExtension, Word, DOMParser */

import { Base64 } from "./base64.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      document.getElementById("error-box").innerHTML +=
        "Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.<br/>";
    }

    // enable extended logging
    OfficeExtension.config.extendedErrorLogging = true;

    // Assign event handlers and other initialization logic.
    document.getElementById("link-hero").onclick = function () {
      changeCitations("hero");
    };
    document.getElementById("link-heronet").onclick = function () {
      changeCitations("heronet");
    };
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("progress-text").innerHTML = "Initializing...";
  }
});

function getURL(url, label) {
  var newURL;
  if (url == "hero") {
    newURL = "https://hero.epa.gov/hero/index.cfm/reference/details/reference_id/" + label;
  }
  if (url == "heronet") {
    newURL = "http://heronet.epa.gov/heronet/index.cfm/reference/details/reference_id/" + label;
  }
  return newURL;
}

function changeURL(url, oldURL) {
  var newURL;
  var heroURL = "https://hero.epa.gov/hero/index.cfm/reference/details/reference_id/";
  var heroNetURL = "http://heronet.epa.gov/heronet/index.cfm/reference/details/reference_id/";
  if (!oldURL.includes(heroURL) && !oldURL.includes(heroNetURL)) {
    return null;
  }
  if (url == "hero") {
    newURL = oldURL.replace(heroNetURL, heroURL);
  }
  if (url == "heronet") {
    newURL = oldURL.replace(heroURL, heroNetURL);
  }
  return newURL;
}

function changeCitations(url) {
  Word.run(function (context) {
    /*
        todo:
        remove check for en.cite.data for decoding
            order doesn't matter, just try and decode all of the fields w/ check
        count uses of each citation in block 2?
            no point in this
        
        things to fix:
            linebreak with issn
            < and > in doi link
            for id 77963, no space after link

        var docBody = context.document.body;
        docBody.insertParagraph("" + decoded,
                                "End");

        */

    // clear error box
    document.getElementById("error-box").innerHTML = '<p style="">Errors and Warnings</p>';
    document.getElementById("loader").style.display = "flex";
    document.getElementById("app-body").style.display = "none";
    document.getElementById("progress-text").innerHTML = "Changing links...";

    var body = context.document.body;
    var bodyOoxml = body.getOoxml();
    var linkRanges = body.getRange("Content").getHyperlinkRanges();
    linkRanges.load("items, hyperlink, font, text");

    var oParser = new DOMParser();
    var bibSearch;
    var bibSearch2;
    var bibSearch3;
    var citationList;
    var citationMatching;
    var bookmarkList;
    var oldBodyXML;

    return context
      .sync()
      .then(function () {
        // search xml to get citations
        var xmlDOM = oParser.parseFromString(bodyOoxml.value, "text/xml");
        citationList = getCitationList(context, xmlDOM); // returns array of citation objects
        bookmarkList = getBibCitations(context, xmlDOM); // returns object with {ENREF: citation text}
        citationMatching = findMatchingCitation(context, citationList, bookmarkList); // object with {ENREF: citation list index}
        assignHeroLinks(context, url, linkRanges, citationMatching, citationList);
        bibSearch = searchForBibText(context, body, bookmarkList); // search for text ranges based on citation text
        document.getElementById("progress-text").innerHTML = "Adding links in bibliography...";

        return context.sync();
      })
      .then(function () {
        // split up text ranges
        bibSearch2 = organizeBibText(context, bibSearch);
        return context.sync();
      })
      .then(function () {
        // find text range to link
        bibSearch3 = findBibTextToLink(context, bibSearch2);
        return context.sync();
      })
      .then(function () {
        // write links to bib ranges
        writeBibLinks(context, url, bibSearch3, citationMatching, citationList);

        var docBody = context.document.body;
        oldBodyXML = docBody.getOoxml();
        document.getElementById("progress-text").innerHTML = "Changing first bibliography entry...";
        return context.sync();
      })
      .then(function () {
        var docBody = context.document.body;
        var newXML = fixProblems(context, url, oParser, oldBodyXML.value, bookmarkList, citationMatching, citationList);
        docBody.insertOoxml(newXML, "Replace");

        // change display for output
        if (document.getElementById("error-box").textContent == "Errors and Warnings") {
          document.getElementById("error-box").innerHTML = '<p style="">Errors and Warnings: None</p>';
        }
        document.getElementById("loader").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("progress-text").innerHTML = "";
        return context.sync();
      });
  }).catch(function (error) {
    document.getElementById("loader").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("error-box").innerHTML += "Fatal error, script exiting...<br/>";
    document.getElementById("error-box").innerHTML += "Error: " + error + "<br/>";
    document.getElementById("progress-text").innerHTML = "Fatal Error";
    if (error instanceof OfficeExtension.Error) {
      document.getElementById("error-box").innerHTML += "Debug info: " + JSON.stringify(error.debugInfo) + "<br/>";
    }
  });
}

function writeString(context, string) {
  // for debugging
  var docBody = context.document.body;
  docBody.insertParagraph("" + string, "End");
}

function fixProblems(context, url, oParser, oldXML, bookmarkList, citationMatching, citationList) {
  // fix a few problems with the code
  var newXML;

  // edit the xml to change the first bibliography entry
  newXML = changeFirstBibEntry(context, url, oParser, oldXML, bookmarkList, citationMatching, citationList);
  if (newXML !== null) {
    // remove bookmarks
    for (var ref in bookmarkList) {
      if (!Object.prototype.hasOwnProperty.call(bookmarkList, ref)) {
        continue;
      }
      var removeBookmark = new RegExp('<w:bookmarkStart w:id="\\d+" w:name="' + ref.replace("#", "") + '"/>', "g");
      newXML = newXML.replace(removeBookmark, "");
    }
    // newXML = newXML.replace(/<w:bookmarkStart w:id="\d+" w:name="_ENREF_\d+"\/>/g, "");
    // newXML = newXML.replace(/<w:bookmarkEnd w:id="\d+"\/>/g, "");
  } else {
    newXML = oldXML;
  }

  // replace carriage returns
  // newXML = newXML.replace(/&amp;#xD;/g, " / ");

  return newXML;
}

function changeFirstBibEntry(context, url, oParser, oldXML, bookmarkList, citationMatching, citationList) {
  // change the XML of the first bib entry to add a link
  // for some reason, the first bib entry isn't included in any one paragraph

  var thisRef = "_ENREF_1";
  if (!oldXML.includes(thisRef)) {
    return null;
  }
  if ("#" + thisRef in bookmarkList) {
    var newURL = getURL(url, citationList[citationMatching["#" + thisRef]].label);
    var citationText = bookmarkList["#" + thisRef];
  } else {
    return null;
  }

  // parse citation string, copied from earlier in script
  var citeSplit = citationText.split(".");

  var runningCount = 0;
  var parenCount = 0;
  for (var j = 0; j < citeSplit.length; j++) {
    var searchMatch = citeSplit[j];
    if (searchMatch.includes("(")) {
      parenCount = parenCount + 1;
    }
    if (searchMatch.includes(")")) {
      parenCount = parenCount - 1;
    }
    runningCount = runningCount + searchMatch.length;
    if (searchMatch.length < 4) {
      continue;
    }
    if (searchMatch.match(/[(][^)]{4,}[a-z]?[)][.]?/)) {
      break;
    }
    if (j > 4) {
      break;
    }
    if (runningCount > 40 && parenCount == 0) {
      break;
    }
  }

  var combRange = "";
  if (runningCount > 0) {
    for (var q = 0; q < j + 1; q++) {
      var delim = "";
      if (q < citeSplit.length - 1) {
        delim = ".";
      }
      combRange = combRange + citeSplit[q] + delim;
    }
  }
  if (combRange.length == 0) {
    document.getElementById("error-box").innerHTML +=
      '<p class="p-warn"><span class="style-err">' +
      "Could not find any text to link in the first citation.</span></p>";
    return null;
  }

  // find section to change
  var finalXMLstr = "";
  var xmlDOM = oParser.parseFromString(oldXML, "text/xml");
  var bookmarkList2 = xmlDOM.getElementsByTagName("w:bookmarkStart");
  for (var a = 0; a < bookmarkList2.length; a++) {
    var bookmark = bookmarkList2[a];
    if (bookmark.hasAttribute("w:name") && bookmark.getAttribute("w:name") == "_ENREF_1") {
      // get correct element, get parent, insert text in right place
      var thisParagraph = bookmark.parentNode;
      if (thisParagraph.nodeName != "w:p" || !thisParagraph.hasAttribute("w:rsidRPr")) {
        continue;
      }
      var rsid = thisParagraph.getAttribute("w:rsidRPr");

      var newText =
        "" +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> HYPERLINK "' +
        newURL +
        '" </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r w:rsidRPr="' +
        rsid +
        '"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
        "<w:t>" +
        combRange +
        "</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

      var splitMatch = oldXML.match(/<w:bookmarkStart w:id="\d" w:name="_ENREF_1"\/>/);
      if (!splitMatch) {
        document.getElementById("error-box").innerHTML += "OH NO";
        continue;
      }
      var splitStr = splitMatch[0];
      var splitArr = oldXML.split(splitStr);
      finalXMLstr = splitArr[0] + splitStr + newText + splitArr[1];

      break;
    } else {
      continue;
    }
  }

  // document.getElementById("error-box").innerHTML += citationText + "<br><br>";

  if (finalXMLstr.length == 0) {
    document.getElementById("error-box").innerHTML +=
      '<p class="p-warn"><span class="style-err">' + "Could not add a link to the first bibliography entry.</span></p>";
    return null;
  }

  var xmlStringedit = finalXMLstr.replace(citationText, citationText.replace(combRange, ""));

  return xmlStringedit;
}

function writeBibLinks(context, url, bibSearch3, citationMatching, citationList) {
  // write hyperlinks onto ranges
  for (var ref in bibSearch3) {
    if (!Object.prototype.hasOwnProperty.call(bibSearch3, ref)) {
      continue;
    }
    var label = citationList[citationMatching[ref]].label;
    var searchList = bibSearch3[ref];
    for (var i = 0; i < searchList.length; i++) {
      var searchCollection = searchList[i];

      searchCollection.hyperlink = getURL(url, label);
    }
  }
}

function findBibTextToLink(context, bibSearch2) {
  // find ranges to change into hyperlinks
  var bibSearch3 = new Object();
  for (var ref in bibSearch2) {
    if (!Object.prototype.hasOwnProperty.call(bibSearch2, ref)) {
      continue;
    }
    var searchCollectionList = bibSearch2[ref];

    var searchList = [];
    for (var i = 0; i < searchCollectionList.length; i++) {
      // should only have length of 1
      var searchCollection = searchCollectionList[i];
      var runningCount = 0;
      var parenCount = 0;
      for (var j = 0; j < searchCollection.items.length; j++) {
        var searchMatch = searchCollection.items[j];
        if (searchMatch.text.includes("(")) {
          parenCount = parenCount + 1;
        }
        if (searchMatch.text.includes(")")) {
          parenCount = parenCount - 1;
        }
        runningCount = runningCount + searchMatch.text.length;
        if (searchMatch.text.length < 4) {
          continue;
        }
        if (searchMatch.text.match(/[(][^)]{4,}[)][.]/)) {
          break;
        }
        if (j > 4) {
          break;
        }
        if (runningCount > 40 && parenCount == 0) {
          break;
        }
      }

      if (runningCount > 0) {
        var combRange = searchCollection.items[0].getRange().expandTo(searchCollection.items[j].getRange()).getRange();
        combRange.load("text, hyperlink");
        searchList.push(combRange);
      } else {
        document.getElementById("error-box").innerHTML +=
          '<p class="p-warn"><span class="style-err">' +
          "Could find text to add ref " +
          ref +
          " to a bibliography entry.</span></p>";
      }
    }
    bibSearch3[ref] = searchList;
  }
  return bibSearch3;
}

function organizeBibText(context, bibSearch) {
  // split the search results into text ranges
  var bibSearch2 = new Object();
  for (var ref in bibSearch) {
    if (!Object.prototype.hasOwnProperty.call(bibSearch, ref)) {
      continue;
    }
    var searchList = [];
    var searchCollection = bibSearch[ref];
    for (var j = 0; j < searchCollection.items.length; j++) {
      var searchMatch = searchCollection.items[j];
      var searchResults = searchMatch.getTextRanges(["."]);
      searchResults.load("text");
      searchList.push(searchResults);
    }
    bibSearch2[ref] = searchList;
  }
  return bibSearch2;
}

function searchForBibText(context, body, bookmarkList) {
  /*
    now we have an object with {label: citation text}
    we need to search the text for (citation text), get the part to underline, and assign a hyperlink
    */
  var searchList = new Object();
  for (var ref in bookmarkList) {
    if (!Object.prototype.hasOwnProperty.call(bookmarkList, ref)) {
      continue;
    }
    if (ref == "#_ENREF_1") {
      continue;
    }
    var searchResults = body.search(bookmarkList[ref].substring(0, 255), { matchCase: true });
    searchResults.load("items, text");
    searchList[ref] = searchResults;
  }
  return searchList;
}

function assignHeroLinks(context, url, linkRanges, citationMatching, citationList) {
  // change links in document
  for (var n = 0; n < linkRanges.items.length; n++) {
    var oldURL = linkRanges.items[n].hyperlink;
    var oldText = linkRanges.items[n].text;

    if (oldURL in citationMatching && oldURL != oldText) {
      var newURL = getURL(url, citationList[citationMatching[oldURL]].label);
      linkRanges.items[n].hyperlink = newURL;
    } else {
      var changedURL = changeURL(url, oldURL);
      if (changedURL !== null) {
        if (oldText == oldURL) {
          linkRanges.items[n].insertText(changedURL, "Replace");
        }
        linkRanges.items[n].hyperlink = changedURL;
      } else if (oldText != oldURL) {
        var errStyle = "info";
        if (oldURL.includes("_ENREF_")) {
          linkRanges.items[n].font.highlightColor = "#FFFF00";
          errStyle = "err";
        }
        document.getElementById("error-box").innerHTML +=
          '<p class="p-warn"><span class="style-' +
          errStyle +
          '">' +
          'Hyperlink "' +
          oldText +
          '" ("' +
          oldURL +
          '")' +
          " not changed.</span></p>";
      }
    }
  }
}

function findMatchingCitation(context, citationList, bookmarkList) {
  // match bib entries to a citation

  // use year/author/title/check other info
  // make an object with enref: label
  // check how many factors match, if less than all, iterate through all citations to check
  // {label: thisLabel, author: thisAuthor, year: thisYear, title: thisTitle}
  // there may be duplicate citations

  var returnValue = new Object();
  var usedCitations = [];
  for (var ref in bookmarkList) {
    if (!Object.prototype.hasOwnProperty.call(bookmarkList, ref)) {
      continue;
    }
    var citeInd = -1;
    var thisBib = bookmarkList[ref];
    var tieBreak = -1;
    for (var a = 0; a < citationList.length; a++) {
      if (usedCitations.includes(a)) {
        continue;
      }
      var thisCitation = citationList[a];
      var checkPass = 0;

      // check for matching text
      if (thisCitation.author != "" && thisBib.includes(thisCitation.author)) {
        checkPass += 2;
      }
      if (thisCitation.year != "" && thisBib.includes(thisCitation.year)) {
        checkPass += 1;
      }
      if (thisCitation.title != "" && thisBib.includes(thisCitation.title)) {
        checkPass += 4;
      }

      if (checkPass == 7) {
        citeInd = a;
        tieBreak = checkPass;
        break;
      }

      if (checkPass < 4) {
        continue;
      }

      if (citeInd != -1) {
        if (checkPass > tieBreak) {
          citeInd = a;
          tieBreak = checkPass;
        } else if (checkPass == tieBreak) {
          // keep current one
        } else {
          // keep current one
        }
      } else {
        citeInd = a;
        tieBreak = checkPass;
      }
    }
    if (citeInd == -1) {
      document.getElementById("error-box").innerHTML +=
        '<p class="p-warn"><span class="style-err">' +
        'Could not find a matching citation for bib entry "' +
        thisBib.split(".")[0] +
        '...".</span></p>';
    } else {
      if (tieBreak < 7) {
        document.getElementById("error-box").innerHTML +=
          '<p class="p-warn"><span class="style-warn">' +
          'Matched citation for bib entry "' +
          thisBib.split(".")[0] +
          '...", but it was not a full match.</span></p>';
      }
      returnValue[ref] = citeInd;
      usedCitations.push(citeInd);
    }
  }
  return returnValue;
}

function getBibCitations(context, xmlDOM) {
  // search for the bibliography text
  // there is a preview api for the bookmark that would be good to switch to

  // var bibXML = OoxmlList[1].value;
  var returnValue = new Object();
  var bookmarkList = xmlDOM.getElementsByTagName("w:bookmarkStart");
  for (var a = 0; a < bookmarkList.length; a++) {
    var bookmark = bookmarkList[a];
    if (bookmark.hasAttribute("w:name")) {
      var anchorName = bookmark.getAttribute("w:name");
      if (!anchorName.includes("_ENREF_")) {
        continue;
      }

      // get text for citation
      var findText = bookmark.nextSibling;
      var foundText = false;
      var citeValue;
      while (!foundText && findText !== null) {
        if (findText.nodeName == "w:r") {
          var textNode = findText.firstChild;
          if (textNode !== null && textNode.nodeName == "w:t") {
            // if (textNode.hasAttribute("xml:space") && textNode.getAttribute("xml:space") == "preserve") {
            var textValue = textNode.textContent;
            if (textValue !== null && textValue.length > 40) {
              foundText = true;
              citeValue = textValue;
              continue;
            }
          }
        }
        findText = findText.nextSibling;
      }
      if (!foundText) {
        continue;
      }

      // set return values
      returnValue["#" + anchorName] = citeValue;
    } else {
      // no bookmark name, skip
      continue;
    }
  }

  return returnValue;
}

function getCitationList(context, xmlDOM) {
  // get list of citations from XML

  var decodeList = [];
  var citationList = [];
  var citationGet = xmlDOM.getElementsByTagName("w:instrText");
  var extraGet = xmlDOM.getElementsByTagName("w:fldData");
  if (extraGet === null) {
    extraGet = [];
  }
  if (extraGet.length > 0) {
    decodeList.push(extraGet[0].textContent);
  }
  var sameCt = 1;
  for (var aa = 1; aa < extraGet.length; aa++) {
    // if (extraGet.length % 2 == 0) {
    //     if (aa % 2 == 0) {
    //         decodeList.push(extraGet[aa].textContent);
    //     }
    // }
    if (extraGet[aa - 1].textContent != extraGet[aa].textContent || sameCt > 1) {
      decodeList.push(extraGet[aa].textContent);
      sameCt = 0;
    } else {
      sameCt += 1;
    }
  }
  var decodeCt = 0;
  for (var jj = 0; jj < citationGet.length; jj++) {
    var tempContent = citationGet[jj].textContent;
    if (tempContent.includes(" ADDIN EN.CITE ") && tempContent != " ADDIN EN.CITE ") {
      citationList.push(tempContent);
    }
    if (extraGet.length > 0 && tempContent == " ADDIN EN.CITE.DATA ") {
      citationList.push("DECODE");
      decodeCt += 1;
    }
  }

  // this checks to make sure the number of EN.CITE.DATA matches the number of w:fldData tags
  var canDecode = true;
  if (decodeCt != decodeList.length) {
    document.getElementById("error-box").innerHTML +=
      '<p class="p-warn"><span class="style-err">' +
      "Error decoding citations; could not citations to fields.</span></p>";
    canDecode = false;
  }

  // decode data
  var newCitationText = []; // list of citation objects
  var citeCt = 0;
  for (var dd = 0; dd < citationList.length; dd++) {
    // decode citations and make object for each with text data
    var decoded;
    if (citationList[dd] == "DECODE") {
      decoded = "";
      if (canDecode) {
        var decodeText = decodeList[citeCt].split("\n"); // string of base64
        for (var tt = 0; tt < decodeText.length; tt++) {
          var thisText = decodeText[tt].replace("\r", "");
          if (thisText.length > 0) {
            decoded += Base64.decode(thisText);
          }
        }
      }
      // decoded = decodeXml(decoded);
      citeCt += 1;
    } else {
      decoded = citationList[dd];
    }

    if (decoded.length > 0) {
      // make citation objects
      newCitationText = fixCombinedCitations(context, newCitationText, decoded);
    }
  }
  return newCitationText;
}

function decodeXml(string) {
  // derived from https://stackoverflow.com/questions/7918868
  return string.replace(/(&quot;|&lt;|&gt;|&amp;|&apos;)/g, function (c) {
    switch (c) {
      case "&lt;":
        return "<";
      case "&gt;":
        return ">";
      case "&amp;":
        return "&";
      case "&apos;":
        return "'";
      case "&quot;":
        return '"';
    }
  });
}

function fixCombinedCitations(context, newCitationText, decoded) {
  // pull citation data out of combined citations

  // check to see if the text is indeed a citation
  if (!decoded.includes("<EndNote><Cite>")) {
    return newCitationText;
  }

  var tempList = new Object(); // object for checking for dups
  var matchCitations = decoded.split("<Cite>");
  for (var ww = 0; ww < matchCitations.length; ww++) {
    var thisAuthor = "";
    var thisYear = "";
    var thisLabel = "";
    var thisTitle = "";

    var matchLabelTest = matchCitations[ww].match(/<label>(\d{1,})<\/label>/g);
    if (matchLabelTest) {
      if (matchLabelTest.length == 1) {
        thisLabel = matchLabelTest[0];
      } else {
        // prevent duplicate labels
        var tempLabel = [];
        for (var cc = 0; cc < matchLabelTest.length; cc++) {
          if (!tempLabel.includes(matchLabelTest[cc])) {
            tempLabel.push(matchLabelTest[cc]);
          }
        }
        if (tempLabel.length == 1) {
          thisLabel = tempLabel[0];
        } else {
          // leave label blank
        }
      }
    }

    // don't add if duplicate or blank
    if (thisLabel == "") {
      continue;
    }
    thisLabel = thisLabel.match(/<label>(\d{1,})<\/label>/)[1];
    // if (labelList.includes(thisLabel)) {continue;}

    var matchAuthor = matchCitations[ww].match(/<Author>([^<>/]{1,})<\/Author>/);
    var matchAuthorTest = matchCitations[ww].match(/<Author>([^<>/]{1,})<\/Author>/g);
    if (matchAuthor && matchAuthorTest && matchAuthorTest.length == 1) {
      thisAuthor = decodeXml(matchAuthor[1]);
    }
    var matchYear = matchCitations[ww].match(/<Year>(\d{4})<\/Year>/);
    var matchYearTest = matchCitations[ww].match(/<Year>(\d{4})<\/Year>/g);
    if (matchYear && matchYearTest && matchYearTest.length == 1) {
      thisYear = matchYear[1];
    }
    var matchTitle = matchCitations[ww].match(/<title>([^<>]{1,})<\/title>/);
    var matchTitleTest = matchCitations[ww].match(/<title>([^<>]{1,})<\/title>/g);
    if (matchTitle && matchTitleTest && matchTitleTest.length == 1) {
      thisTitle = decodeXml(matchTitle[1]);
    }

    // check for duplicate citations
    var combData = thisAuthor + thisYear + thisTitle;
    if (thisLabel in tempList) {
      if (tempList[thisLabel].includes(combData)) {
        continue;
      } else {
        tempList[thisLabel].push(combData);
      }
    } else {
      tempList[thisLabel] = [combData];
    }

    var newCitation = { label: thisLabel, author: thisAuthor, year: thisYear, title: thisTitle };
    newCitationText.push(newCitation);
  }
  return newCitationText;
}
