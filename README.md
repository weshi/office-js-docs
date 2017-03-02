
# Word JavaScript APIs

Welcome to the Word JavaScript API open specification documentation. The Word JavaScript API provides Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. This branch documents the new APIs that our team is working on and plans to ship in the next few months.

In order to try the Preview APIs please make sure to: 

1. Get on the First Release so you can get the latest Word! Follow the instructions here: https://github.com/OfficeDev/office-js-docs/blob/215f5d35490c943cc06c29b98357ba8cb034ec81/docs/develop/install-latest-office-version.md
2. Use the PREVEW CDN for OFfice.js:https://appsforoffice.microsoft.com/lib/beta/hosted/office.js
3. Code and provide feedback!!!


## Introduction to Word JavaScript API 1.3 

This section describes the new set of Word JavaScript APIs that are planned for the next release (requirement set 1.3). To provide feedback on the new APIs, use the links in the following table to open issues in GitHub. 

_**Note**: The features listed are still in the design and review phase and are not yet available as part of the product. The final design is subject to change. When the feature is made available, the final specification will be published as part of the master repository._

### FEATURES
|Object| What is new| Description|Req. Set|
|:----|:----|:----|:----|
|[application](reference/word/application.md)|_Method_ > [createDocument(base64File: string)](application.md#createdocumentbase64file-string)|Creates a new document by using a base64 encoded .docx file.|1.4|
|[document](reference/word/document.md)|_Property_ > settings|Gets the add-in's settings in the current document. Read-only.|1.4|
|[document](reference/word/document.md)|_Method_ > [deleteBookmark(name: string)](document.md#deletebookmarkname-string)|Deletes a bookmark, if exists, from the document.|1.4|
|[document](reference/word/document.md)|_Method_ > [getBookmarkRange(name: string)](document.md#getbookmarkrangename-string)|Gets a bookmark's range. Throws if the bookmark does not exist.|1.4|
|[document](reference/word/document.md)|_Method_ > [getBookmarkRangeOrNullObject(name: string)](document.md#getbookmarkrangeornullobjectname-string)|Gets a bookmark's range. Returns a null object if the bookmark does not exist.|1.4|
|[document](reference/word/document.md)|_Method_ > [open()](document.md#open)|Open the document.|1.4|
|[inlinePicture](reference/word/inlinepicture.md)|_Property_ > imageFormat|Gets the format of the inline image. Read-only. Possible values are: Unsupported, Undefined, Bmp, Jpeg, Gif, Tiff, Png, Icon, Exif, Wmf, Emf, Pict, Pdf, Svg.|1.4|
|[list](reference/word/list.md)|_Method_ > [getLevelFont(level: number)](list.md#getlevelfontlevel-number)|Gets the font of the bullet, number or picture at the specified level in the list.|1.4|
|[list](reference/word/list.md)|_Method_ > [getLevelPicture(level: number)](list.md#getlevelpicturelevel-number)|Gets the base64 encoded string representation of the picture at the specified level in the list.|1.4|
|[list](reference/word/list.md)|_Method_ > [resetLevelFont(level: number, resetFontName: bool)](list.md#resetlevelfontlevel-number-resetfontname-bool)|Resets the font of the bullet, number or picture at the specified level in the list.|1.4|
|[list](reference/word/list.md)|_Method_ > [setLevelPicture(level: number, base64EncodedImage: string)](list.md#setlevelpicturelevel-number-base64encodedimage-string)|Sets the picture at the specified level in the list.|1.4|
|[range](reference/word/range.md)|_Method_ > [getBookmarks(includeHidden: bool, includeAdjacent: bool)](range.md#getbookmarksincludehidden-bool-includeadjacent-bool)|Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character.|1.4|
|[range](reference/word/range.md)|_Method_ > [insertBookmark(name: string)](range.md#insertbookmarkname-string)|Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first.|1.4|
|[section](reference/word/section.md)|_Property_ > headerFooterEvenPageDifferent|Gets or sets a value that indicates whether even-numbered pages have a different header and footer from odd-numbered pages in the section.|1.4|
|[section](reference/word/section.md)|_Property_ > headerFooterFirstPageDifferent|Gets or sets a value that indicates whether the first page has a different header and footer from the other pages in the section.|1.4|
|[setting](reference/word/setting.md)|_Property_ > key|Gets the key of the setting. Read only. Read-only.|1.4|
|[setting](reference/word/setting.md)|_Property_ > value|Gets or sets the value of the setting.|1.4|
|[setting](reference/word/setting.md)|_Method_ > [delete()](setting.md#delete)|Deletes the setting.|1.4|
|[settingCollection](reference/word/settingcollection.md)|_Property_ > items|A collection of setting objects. Read-only.|1.4|
|[settingCollection](reference/word/settingcollection.md)|_Method_ > [deleteAll()](settingcollection.md#deleteall)|Deletes all settings in this add-in.|1.4|
|[settingCollection](reference/word/settingcollection.md)|_Method_ > [getCount()](settingcollection.md#getcount)|Gets the count of settings.|1.4|
|[settingCollection](reference/word/settingcollection.md)|_Method_ > [getItem(key: string)](settingcollection.md#getitemkey-string)|Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist.|1.4|
|[settingCollection](reference/word/settingcollection.md)|_Method_ > [getItemOrNullObject(key: string)](settingcollection.md#getitemornullobjectkey-string)|Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.|1.4|
|[settingCollection](reference/word/settingcollection.md)|_Method_ > [add(key: string, value: object)](settingcollection.md#setkey-string-value-object)|Creates a new setting entry.|1.4|
|[table](reference/word/table.md)|_Method_ > [mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](table.md#mergecellstoprow-number-firstcell-number-bottomrow-number-lastcell-number)|Merges the cells bounded inclusively by a first and last cell.|1.4|
|[tableCell](reference/word/tablecell.md)|_Method_ > [split(rowCount: number, columnCount: number)](tablecell.md#splitrowcount-number-columncount-number)|Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.|1.4|
|[tableRow](reference/word/tablerow.md)|_Method_ > [merge()](tablerow.md#merge)|Merges the row into one cell.|1.4|


## Try it out 

_**Note**: This feature is not currently available for new features._


We've been working on a Snippet Explorer that you can use to browse through common code snippets and learn how the new APIs work. Give it a try. You can find the code snippets referenced by the Snippet Explorer at [Office 2016 JavaScript API Snippet Explorer (Preview)](https://officesnippetexplorer.azurewebsites.net/#/snippets/word). 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your feedback by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository. 

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js] and [word].
