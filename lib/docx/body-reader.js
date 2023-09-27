exports.createBodyReader = createBodyReader;
exports._readNumberingProperties = readNumberingProperties;

var dingbatToUnicode = require("dingbat-to-unicode");
var _ = require("underscore");

var documents = require("../documents");
var Result = require("../results").Result;
var warning = require("../results").warning;
var uris = require("./uris");

function createBodyReader(options) {
  return {
    readXmlElement: function (element) {
      return new BodyReader(options).readXmlElement(element);
    },
    readXmlElements: function (elements) {
      return new BodyReader(options).readXmlElements(elements);
    },
  };
}

function BodyReader(options) {
  var complexFieldStack = [];
  var currentInstrText = [];
  var relationships = options.relationships;
  var contentTypes = options.contentTypes;
  var docxFile = options.docxFile;
  var files = options.files;
  var numbering = options.numbering;
  var styles = options.styles;

  function readXmlElements(elements) {
    var results = elements.map(readXmlElement);
    return combineResults(results);
  }

  function readXmlElement(element) {
    if (element.type === "element") {
      var handler = xmlElementReaders[element.name];
      if (handler) {
        return handler(element);
      } else if (
        !Object.prototype.hasOwnProperty.call(ignoreElements, element.name)
      ) {
        var message = warning(
          "An unrecognised element was ignored: " + element.name
        );
        return emptyResultWithMessages([message]);
      }
    }
    return emptyResult();
  }

  function readParagraphIndent(element) {
    return {
      start: element.attributes["w:start"] || element.attributes["w:left"],
      end: element.attributes["w:end"] || element.attributes["w:right"],
      firstLine: element.attributes["w:firstLine"],
      hanging: element.attributes["w:hanging"],
    };
  }

  // CHANGE HERE - added reading borders
  function readParagraphBorders(element) {
    if(!element) {
      return {};
    }
    const top = element.firstOrEmpty("w:top");
    const bottom = element.firstOrEmpty("w:bottom");
    const left = element.firstOrEmpty("w:left");
    const right = element.firstOrEmpty("w:right");
    return {
      top,bottom,left,right
    }
  }

  // CHANGE HERE - added spacing
  function readParagraphSpacing(element) {
    if (!element.attributes["w:before"] && !element.attributes["w:after"]) {
      return {};
    }
    function spacingToPixels(spacing) {
      let points = spacing / 20;
      let pixels = points * (96 / 72);
      return pixels;
    }
    const before = parseInt(element.attributes["w:before"] || "0", 10);
    const after = parseInt(element.attributes["w:after"] || "0", 10);
    return {
      before: element.attributes["w:before"] ? spacingToPixels(before) : null,
      after: element.attributes["w:after"] ? spacingToPixels(after) : null,
    };
  }

  function readRunProperties(element) {
    return readRunStyle(element).map(function (style) {
      var fontSizeString = element.firstOrEmpty("w:sz").attributes["w:val"];
      // w:sz gives the font size in half points, so halve the value to get the size in points
      var fontSize = /^[0-9]+$/.test(fontSizeString)
        ? parseInt(fontSizeString, 10) / 2
        : null;

      // CHANGE HERE - added a check for w:szCs as an alternative to w:sz
      // note. that there could also be http://officeopenxml.com/WPtextFormatting.php szCs for complex languages
      // in fact, we can see it in some other examples too so worth checking for it
      if (!fontSize) {
        fontSizeString = element.firstOrEmpty("w:szCs").attributes["w:val"];
        fontSize = /^[0-9]+$/.test(fontSizeString)
          ? parseInt(fontSizeString, 10) / 2
          : null;
      }

      return {
        type: "runProperties",
        styleId: style.styleId,
        styleName: style.name,
        verticalAlignment:
          element.firstOrEmpty("w:vertAlign").attributes["w:val"],
        font: element.firstOrEmpty("w:rFonts").attributes["w:ascii"],
        fontSize: fontSize,
        isBold: readBooleanElement(element.first("w:b")),
        isUnderline: readUnderline(element.first("w:u")),
        isItalic: readBooleanElement(element.first("w:i")),
        isStrikethrough: readBooleanElement(element.first("w:strike")),
        isAllCaps: readBooleanElement(element.first("w:caps")),
        isSmallCaps: readBooleanElement(element.first("w:smallCaps")),

        // CHANGE HERE to surface the color property
        fontColor: readColor(element.first("w:color")),
      };
    });
  }

  // CHANGE HERE to surface the color property
  function readColor(element) {
    if (element) {
      var value = element.attributes["w:val"];
      if(!value){
        return null
      }
      // CHANGE HERE - force auto to be black
      if(value === 'auto'){
        return '#000000'
      }
      return `#${value}`;
    } else {
      return undefined;
    }
  }

  function readUnderline(element) {
    if (element) {
      var value = element.attributes["w:val"];
      return (
        value !== undefined &&
        value !== "false" &&
        value !== "0" &&
        value !== "none"
      );
    } else {
      // CHANGE HERE - we should allow null values
      return undefined;
    }
  }

  function readBooleanElement(element) {
    if (element) {
      var value = element.attributes["w:val"];
      return value !== "false" && value !== "0" && value!=='none';
    } else {
      // CHANGE HERE - we should allow null values for isBold etc
      // as that's totally legit
      return undefined;
    }
  }

  function readParagraphStyle(element) {
    return readStyle(
      element,
      "w:pStyle",
      "Paragraph",
      styles.findParagraphStyleById
    );
  }

  function readRunStyle(element) {
    return readStyle(element, "w:rStyle", "Run", styles.findCharacterStyleById);
  }

  function readTableStyle(element) {
    return readStyle(element, "w:tblStyle", "Table", styles.findTableStyleById);
  }

  function readStyle(element, styleTagName, styleType, findStyleById) {
    var messages = [];
    var styleElement = element.first(styleTagName);
    var styleId = null;
    var name = null;
    if (styleElement) {
      styleId = styleElement.attributes["w:val"];
      if (styleId) {
        var style = findStyleById(styleId);
        if (style) {
          name = style.name;
        } else {
          messages.push(undefinedStyleWarning(styleType, styleId));
        }
      }
    }
    return elementResultWithMessages(
      { styleId: styleId, name: name },
      messages
    );
  }

  var unknownComplexField = { type: "unknown" };

  function readFldChar(element) {
    var type = element.attributes["w:fldCharType"];
    if (type === "begin") {
      complexFieldStack.push(unknownComplexField);
      currentInstrText = [];
    } else if (type === "end") {
      complexFieldStack.pop();
    } else if (type === "separate") {
      var hyperlinkOptions = parseHyperlinkFieldCode(currentInstrText.join(""));
      var complexField =
        hyperlinkOptions === null
          ? unknownComplexField
          : { type: "hyperlink", options: hyperlinkOptions };
      complexFieldStack.pop();
      complexFieldStack.push(complexField);
    }
    return emptyResult();
  }

  function currentHyperlinkOptions() {
    var topHyperlink = _.last(
      complexFieldStack.filter(function (complexField) {
        return complexField.type === "hyperlink";
      })
    );
    return topHyperlink ? topHyperlink.options : null;
  }

  function parseHyperlinkFieldCode(code) {
    var externalLinkResult = /\s*HYPERLINK "(.*)"/.exec(code);
    if (externalLinkResult) {
      return { href: externalLinkResult[1] };
    }

    var internalLinkResult = /\s*HYPERLINK\s+\\l\s+"(.*)"/.exec(code);
    if (internalLinkResult) {
      return { anchor: internalLinkResult[1] };
    }

    return null;
  }

  function readInstrText(element) {
    currentInstrText.push(element.text());
    return emptyResult();
  }

  function readSymbol(element) {
    // See 17.3.3.30 sym (Symbol Character) of ECMA-376 4th edition Part 1
    var font = element.attributes["w:font"];
    var char = element.attributes["w:char"];
    var unicodeCharacter = dingbatToUnicode.hex(font, char);
    if (unicodeCharacter == null && /^F0..$/.test(char)) {
      unicodeCharacter = dingbatToUnicode.hex(font, char.substring(2));
    }

    if (unicodeCharacter == null) {
      return emptyResultWithMessages([
        warning(
          "A w:sym element with an unsupported character was ignored: char " +
            char +
            " in font " +
            font
        ),
      ]);
    } else {
      return elementResult(new documents.Text(unicodeCharacter.string));
    }
  }

  function noteReferenceReader(noteType) {
    return function (element) {
      var noteId = element.attributes["w:id"];
      return elementResult(
        new documents.NoteReference({
          noteType: noteType,
          noteId: noteId,
        })
      );
    };
  }

  function readCommentReference(element) {
    return elementResult(
      documents.commentReference({
        commentId: element.attributes["w:id"],
      })
    );
  }

  function readChildElements(element) {
    return readXmlElements(element.children);
  }

  var xmlElementReaders = {
    "w:p": function (element) {
      return readXmlElements(element.children)
        .map(function (children) {
          var properties = _.find(children, isParagraphProperties);

          // CHANGE HERE to set the element on the properties of the object
          // to include the element index and absolute index
          if (!properties) {
            properties = {};
          }
          properties.elementIndex = element.elementIndex;
          properties.absoluteIndex = element.absoluteIndex;

          // CHANGE HERE to include lastRenderedPageBreak
          // Note. It will always be at the start of the first run of a paragraph
          properties.hasLastRenderedPageBreak = element.children
            ?.find((child) => child.name === "w:r")
            ?.children?.find(
              (child) => child.name === "w:lastRenderedPageBreak"
            )
            ? true
            : null;

          return new documents.Paragraph(
            children.filter(negate(isParagraphProperties)),
            properties
          );
        })
        .insertExtra();
    },
    "w:pPr": function (element) {
      return readParagraphStyle(element).map(function (style) {
        return {
          type: "paragraphProperties",
          styleId: style.styleId,
          styleName: style.name,
          alignment: element.firstOrEmpty("w:jc").attributes["w:val"],
          numbering: readNumberingProperties(
            style.styleId,
            element.firstOrEmpty("w:numPr"),
            numbering
          ),
          indent: readParagraphIndent(element.firstOrEmpty("w:ind")),

          // CHANGE HERE - adding reading borders
          borders: readParagraphBorders(element.firstOrEmpty("w:pBdr")),

          // CHANGE HERE - added spacing and end of page break
          spacing: readParagraphSpacing(element.firstOrEmpty("w:spacing")),

          // CHANGE HERE - pick out the run properties
          runProperties: readRunProperties(element.firstOrEmpty("w:rPr"))
        };
      });
    },
    "w:r": function (element) {
      return readXmlElements(element.children).map(function (children) {
        var properties = _.find(children, isRunProperties);
        children = children.filter(negate(isRunProperties));

        var hyperlinkOptions = currentHyperlinkOptions();
        if (hyperlinkOptions !== null) {
          children = [new documents.Hyperlink(children, hyperlinkOptions)];
        }

        return new documents.Run(children, properties);
      });
    },
    "w:rPr": readRunProperties,
    "w:fldChar": readFldChar,
    "w:instrText": readInstrText,
    "w:t": function (element) {
      return elementResult(new documents.Text(element.text()));
    },
    "w:tab": function (element) {
      return elementResult(new documents.Tab());
    },
    "w:noBreakHyphen": function () {
      return elementResult(new documents.Text("\u2011"));
    },
    "w:softHyphen": function (element) {
      return elementResult(new documents.Text("\u00AD"));
    },
    "w:sym": readSymbol,
    "w:hyperlink": function (element) {
      var relationshipId = element.attributes["r:id"];
      var anchor = element.attributes["w:anchor"];
      return readXmlElements(element.children).map(function (children) {
        function create(options) {
          var targetFrame = element.attributes["w:tgtFrame"] || null;

          return new documents.Hyperlink(
            children,
            _.extend({ targetFrame: targetFrame }, options)
          );
        }

        if (relationshipId) {
          var href = relationships.findTargetByRelationshipId(relationshipId);
          if (anchor) {
            href = uris.replaceFragment(href, anchor);
          }
          return create({ href: href });
        } else if (anchor) {
          return create({ anchor: anchor });
        } else {
          return children;
        }
      });
    },
    "w:tbl": readTable,
    "w:tr": readTableRow,
    "w:tc": readTableCell,
    "w:footnoteReference": noteReferenceReader("footnote"),
    "w:endnoteReference": noteReferenceReader("endnote"),
    "w:commentReference": readCommentReference,
    "w:br": function (element) {
      var breakType = element.attributes["w:type"];
      if (breakType == null || breakType === "textWrapping") {
        return elementResult(documents.lineBreak);
      } else if (breakType === "page") {
        return elementResult(documents.pageBreak);
      } else if (breakType === "column") {
        return elementResult(documents.columnBreak);
      } else {
        return emptyResultWithMessages([
          warning("Unsupported break type: " + breakType),
        ]);
      }
    },
    "w:bookmarkStart": function (element) {
      var name = element.attributes["w:name"];
      if (name === "_GoBack") {
        return emptyResult();
      } else {
        return elementResult(new documents.BookmarkStart({ name: name }));
      }
    },

    "mc:AlternateContent": function (element) {
      return readChildElements(element.first("mc:Fallback"));
    },
    "w:sdt": function (element) {
      return readXmlElements(element.firstOrEmpty("w:sdtContent").children);
    },
    "w:ins": readChildElements,
    "w:object": readChildElements,
    "w:smartTag": readChildElements,
    "w:drawing": readChildElements,
    "w:pict": function (element) {
      return readChildElements(element).toExtra();
    },
    "v:roundrect": readChildElements,
    "v:shape": readChildElements,
    "v:textbox": readChildElements,
    "w:txbxContent": readChildElements,
    "wp:inline": readDrawingElement,
    "wp:anchor": readDrawingElement,
    "v:imagedata": readImageData,
    "v:group": readChildElements,
    "v:rect": readChildElements,
  };

  return {
    readXmlElement: readXmlElement,
    readXmlElements: readXmlElements,
  };

  function readTable(element) {
    // CHANGE HERE - pulling out the table properties and grid http://officeopenxml.com/WPtableProperties.php
    const tblPr = element.firstOrEmpty("w:tblPr");
    const tblGrid = element.firstOrEmpty("w:tblGrid");
    var propertiesResult = readTableProperties(element.firstOrEmpty("w:tblPr"));
    const tableProperties = readXmlElements(element.children)
      .flatMap(calculateRowSpans)
      .flatMap(function (children) {
        return propertiesResult.map(function (properties) {
          // CHANGE HERE - including the elementIndex and absolute index
          const updatedProperties = {
            ...properties,
            elementIndex: element.elementIndex,
            absoluteIndex: element.absoluteIndex,
            tblPr,
            tblGrid,
          };
          return documents.Table(children, updatedProperties);
        });
      });
    return tableProperties;
  }

  function readTableProperties(element) {
    return readTableStyle(element).map(function (style) {
      return {
        styleId: style.styleId,
        styleName: style.name,
      };
    });
  }

  function readTableRow(element) {
    var properties = element.firstOrEmpty("w:trPr");
    var isHeader = !!properties.first("w:tblHeader");
    // CHANGE HERE - pulling out the table properties and grid http://officeopenxml.com/WPtableProperties.php
    return readXmlElements(element.children).map(function (children) {
      return documents.TableRow(children, { isHeader: isHeader });
    });
  }

  //  CHANGE HERE - no change, but TODO we could pull out the table border properties here, given
  // that they are normally set on the style, this seems a little superfluous
  // for practical purposes, we can leave them as default or from the table style
  // can always reevaluate later
  function readTableCell(element) {
    return readXmlElements(element.children).map(function (children) {
      var properties = element.firstOrEmpty("w:tcPr");

      // CHANGE HERE - included bg color
      var backgroundColor =
        properties.firstOrEmpty("w:shd").attributes["w:fill"];

      var gridSpan = properties.firstOrEmpty("w:gridSpan").attributes["w:val"];
      var colSpan = gridSpan ? parseInt(gridSpan, 10) : 1;

      //  CHANGE HERE - included elementIndex and absoluteIndex
      var cell = documents.TableCell(children, {
        colSpan: colSpan,
        elementIndex: element.elementIndex,
        absoluteIndex: element.absoluteIndex,
        backgroundColor,
      });
      cell._vMerge = readVMerge(properties);

      return cell;
    });
  }

  function readVMerge(properties) {
    var element = properties.first("w:vMerge");
    if (element) {
      var val = element.attributes["w:val"];
      return val === "continue" || !val;
    } else {
      return null;
    }
  }

  function calculateRowSpans(rows) {
    var unexpectedNonRows = _.any(rows, function (row) {
      return row.type !== documents.types.tableRow;
    });
    if (unexpectedNonRows) {
      return elementResultWithMessages(rows, [
        warning(
          "unexpected non-row element in table, cell merging may be incorrect"
        ),
      ]);
    }
    var unexpectedNonCells = _.any(rows, function (row) {
      return _.any(row.children, function (cell) {
        return cell.type !== documents.types.tableCell;
      });
    });
    if (unexpectedNonCells) {
      return elementResultWithMessages(rows, [
        warning(
          "unexpected non-cell element in table row, cell merging may be incorrect"
        ),
      ]);
    }

    var columns = {};

    rows.forEach(function (row) {
      var cellIndex = 0;
      row.children.forEach(function (cell) {
        if (cell._vMerge && columns[cellIndex]) {
          columns[cellIndex].rowSpan++;
        } else {
          columns[cellIndex] = cell;
          cell._vMerge = false;
        }
        cellIndex += cell.colSpan;
      });
    });

    rows.forEach(function (row) {
      row.children = row.children.filter(function (cell) {
        return !cell._vMerge;
      });
      row.children.forEach(function (cell) {
        delete cell._vMerge;
      });
    });

    return elementResult(rows);
  }

  function readDrawingElement(element) {
    var blips = element
      .getElementsByTagName("a:graphic")
      .getElementsByTagName("a:graphicData")
      .getElementsByTagName("pic:pic")
      .getElementsByTagName("pic:blipFill")
      .getElementsByTagName("a:blip");

    // CHANGE HERE - if there's no image file - might be graphic. Post back
    // Unknown
    if(blips.length === 0) {
      var image = documents.Image({
        readImage: ()=>{},
        altText: 'Unkown image',
        contentType: 'unknown',
        dimensions: {},
      });
      return elementResultWithMessages(image, []);
    }

    return combineResults(blips.map(readBlip.bind(null, element)));
  }

  function readBlip(element, blip) {
    var properties = element.first("wp:docPr").attributes;
    var dimensions = element.first("wp:extent"); // CHANGE HERE, get the extent for the image size
    var altText = isBlank(properties.descr)
      ? properties.title
      : properties.descr;
    var blipImageFile = findBlipImageFile(blip);
    if (blipImageFile === null) {
      return emptyResultWithMessages([
        warning("Could not find image file for a:blip element"),
      ]);
    } else {
      // CHANGE HERE, pass dimensions
      return readImage(blipImageFile, altText, dimensions);
    }
  }

  function isBlank(value) {
    return value == null || /^\s*$/.test(value);
  }

  function findBlipImageFile(blip) {
    var embedRelationshipId = blip.attributes["r:embed"];
    var linkRelationshipId = blip.attributes["r:link"];
    if (embedRelationshipId) {
      return findEmbeddedImageFile(embedRelationshipId);
    } else if (linkRelationshipId) {
      var imagePath =
        relationships.findTargetByRelationshipId(linkRelationshipId);
      return {
        path: imagePath,
        read: files.read.bind(files, imagePath),
      };
    } else {
      return null;
    }
  }

  function readImageData(element) {
    var relationshipId = element.attributes["r:id"];

    if (relationshipId) {
      return readImage(
        findEmbeddedImageFile(relationshipId),
        element.attributes["o:title"]
      );
    } else {
      return emptyResultWithMessages([
        warning("A v:imagedata element without a relationship ID was ignored"),
      ]);
    }
  }

  function findEmbeddedImageFile(relationshipId) {
    var path = uris.uriToZipEntryName(
      "word",
      relationships.findTargetByRelationshipId(relationshipId)
    );
    return {
      path: path,
      read: docxFile.read.bind(docxFile, path),
    };
  }

  // CHANGE HERE - pass the dimensions
  function readImage(imageFile, altText, dimensions) {
    var contentType = contentTypes.findContentType(imageFile.path);

    var image = documents.Image({
      readImage: imageFile.read,
      altText: altText,
      contentType: contentType,
      dimensions: dimensions, // CHANGE HERE - pass the dimensions
    });
    var warnings = supportedImageTypes[contentType]
      ? []
      : warning(
          "Image of type " +
            contentType +
            " is unlikely to display in web browsers"
        );
    return elementResultWithMessages(image, warnings);
  }

  function undefinedStyleWarning(type, styleId) {
    return warning(
      type +
        " style with ID " +
        styleId +
        " was referenced but not defined in the document"
    );
  }
}

function readNumberingProperties(styleId, element, numbering) {
  // CHANGE - we need to see if there is a local instance of the numId to use
  // if so, use that, otherwise see if there's a styleId to check out as the numId might
  // in there (in that case, we find the numbering Id in the client by traversing the styles object)
  // It shold be noted that it's possible for there to be a numId of 0 or some number
  // that does not exist in the numbering.xml file - in this case, there should be no numbering returned
  var level = element.firstOrEmpty("w:ilvl").attributes["w:val"];
  var numId = element.firstOrEmpty("w:numId").attributes["w:val"];

  // TODO deal with numbering according to styles
  if (numId === undefined && styleId != null) {
    var levelByStyleId = numbering.findLevelByParagraphStyleId(styleId);
    if (levelByStyleId != null) {
      return levelByStyleId;
    }
  }

  if (level === undefined || numId === undefined) {
    return null;
  } else {
    // CHANGE - get an array of the start levels as they are needed
    // for odd cases of restarts - and can be different per level anyway
    const allNumberingForLevel = numbering.getAllNumberingForLevel(numId)
    return {
      // CHANGE - we need to include the 'instance' Id of the numbering
      numberingId: numId,
      allNumberingForLevel,
      ...numbering.findLevel(numId, level),
    };
  }
}

var supportedImageTypes = {
  "image/png": true,
  "image/gif": true,
  "image/jpeg": true,
  "image/svg+xml": true,
  "image/tiff": true,
};

var ignoreElements = {
  "office-word:wrap": true,
  "v:shadow": true,
  "v:shapetype": true,
  "w:annotationRef": true,
  "w:bookmarkEnd": true,
  "w:sectPr": true,
  "w:proofErr": true,
  "w:lastRenderedPageBreak": true,
  "w:commentRangeStart": true,
  "w:commentRangeEnd": true,
  "w:del": true,
  "w:footnoteRef": true,
  "w:endnoteRef": true,
  "w:tblPr": true,
  "w:tblGrid": true,
  "w:trPr": true,
  "w:tcPr": true,
};

function isParagraphProperties(element) {
  return element.type === "paragraphProperties";
}

function isRunProperties(element) {
  return element.type === "runProperties";
}

function negate(predicate) {
  return function (value) {
    return !predicate(value);
  };
}

function emptyResultWithMessages(messages) {
  return new ReadResult(null, null, messages);
}

function emptyResult() {
  return new ReadResult(null);
}

function elementResult(element) {
  return new ReadResult(element);
}

function elementResultWithMessages(element, messages) {
  return new ReadResult(element, null, messages);
}

function ReadResult(element, extra, messages) {
  this.value = element || [];
  this.extra = extra;
  this._result = new Result(
    {
      element: this.value,
      extra: extra,
    },
    messages
  );
  this.messages = this._result.messages;
}

ReadResult.prototype.toExtra = function () {
  return new ReadResult(
    null,
    joinElements(this.extra, this.value),
    this.messages
  );
};

ReadResult.prototype.insertExtra = function () {
  var extra = this.extra;
  if (extra && extra.length) {
    return new ReadResult(joinElements(this.value, extra), null, this.messages);
  } else {
    return this;
  }
};

ReadResult.prototype.map = function (func) {
  var result = this._result.map(function (value) {
    return func(value.element);
  });
  return new ReadResult(result.value, this.extra, result.messages);
};

ReadResult.prototype.flatMap = function (func) {
  var result = this._result.flatMap(function (value) {
    return func(value.element)._result;
  });
  return new ReadResult(
    result.value.element,
    joinElements(this.extra, result.value.extra),
    result.messages
  );
};

function combineResults(results) {
  var result = Result.combine(_.pluck(results, "_result"));
  return new ReadResult(
    _.flatten(_.pluck(result.value, "element")),
    _.filter(_.flatten(_.pluck(result.value, "extra")), identity),
    result.messages
  );
}

function joinElements(first, second) {
  return _.flatten([first, second]);
}

function identity(value) {
  return value;
}
