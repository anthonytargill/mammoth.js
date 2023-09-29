var _ = require("underscore");

exports.readNumberingXml = readNumberingXml;
exports.Numbering = Numbering;
exports.defaultNumbering = new Numbering({}, {});

function Numbering(nums, abstractNums, styles) {
    var allLevels = _.flatten(_.values(abstractNums).map(function(abstractNum) {
        return _.values(abstractNum.levels);
    }));

    var levelsByParagraphStyleId = _.indexBy(
        allLevels.filter(function(level) {
            return level.paragraphStyleId != null;
        }),
        "paragraphStyleId"
    );

    function findLevel(numId, level) {
        var num = nums[numId];
        if (num) {
            var abstractNum = abstractNums[num.abstractNumId];
            if (!abstractNum) {
                return null;
            } else if (abstractNum.numStyleLink == null) {
                return abstractNums[num.abstractNumId].levels[level];
            } else {
                var style = styles.findNumberingStyleById(abstractNum.numStyleLink);
                return findLevel(style.numId, level);
            }
        } else {
            return null;
        }
    }

    function findLevelByParagraphStyleId(styleId) {
        return levelsByParagraphStyleId[styleId] || null;
    }

    // CHANGE HERE - need to extract the numbering so stitch together later
    function getNumbering(){
        return {
            nums, abstractNums, styles
        }
    }

    return {
        getNumbering,
        findLevel: findLevel,
        findLevelByParagraphStyleId: findLevelByParagraphStyleId
    };
}

function readNumberingXml(root, options) {
    if (!options || !options.styles) {
        throw new Error("styles is missing");
    }

    var abstractNums = readAbstractNums(root);
    var nums = readNums(root, abstractNums);
    return new Numbering(nums, abstractNums, options.styles);
}

function readAbstractNums(root) {
    var abstractNums = {};
    root.getElementsByTagName("w:abstractNum").forEach(function(element) {
        var id = element.attributes["w:abstractNumId"];
        abstractNums[id] = readAbstractNum(element);
    });
    return abstractNums;
}

function readAbstractNum(element) {
    var levels = {};
    // CHANGE HERE, build up the start indexes
    let startByLevelMap = {}
    element.getElementsByTagName("w:lvl").forEach(function(levelElement) {
        var levelIndex = levelElement.attributes["w:ilvl"];
        var numFmt = levelElement.first("w:numFmt").attributes["w:val"];
        var paragraphStyleId = levelElement.firstOrEmpty("w:pStyle").attributes["w:val"];

        // CHANGE adding the level text, restart behaviour and isLgl
        // http://officeopenxml.com/WPnumbering-numFmt.php
        // http://officeopenxml.com/WPnumbering-restart.php
        // http://officeopenxml.com/WPnumbering-isLgl.php
        var levelText = levelElement.firstOrEmpty("w:lvlText").attributes["w:val"];
        var levelRestart = levelElement.firstOrEmpty("w:lvlRestart").attributes["w:val"];
        var start = levelElement.firstOrEmpty("w:start").attributes["w:val"];
        var overrideAsNumerals = levelElement.first("w:isLgl")

        // CHANGE here, sometimes the indent is specified in the numbering file
        // per level. In this case, pull it put and add it to the levels text
        var indent = levelElement.firstOrEmpty('w:pPr').firstOrEmpty('w:ind')

        // Add the start index of all of the levels against this one. The reason is that
        // If there isn't a current numbering for that level when it starts, it needs to 
        // backfill the numbering index map with the earlier levels. Only need levels up 
        // to and including this one
        startByLevelMap[levelIndex] = {
            start,
            numberFormat: numFmt
        }
        
        levels[levelIndex] = {
            isOrdered: numFmt !== "bullet",
            level: levelIndex,
            paragraphStyleId: paragraphStyleId,
            indent,

            // CHANGE adding the number format and level text to the output and levelMap
            startByLevelMap,
            numberFormat: numFmt,
            levelText: levelText,
            levelRestart: levelRestart,
            overrideAsNumerals: overrideAsNumerals ? true : false,
            start: start
        };
    });

    var numStyleLink = element.firstOrEmpty("w:numStyleLink").attributes["w:val"];

    return {levels: levels, numStyleLink: numStyleLink};
}

function readNums(root) {
    var nums = {};
    root.getElementsByTagName("w:num").forEach(function(element) {
        var numId = element.attributes["w:numId"];
        var abstractNumId = element.first("w:abstractNumId").attributes["w:val"];
        nums[numId] = {abstractNumId: abstractNumId};
    });
    return nums;
}
