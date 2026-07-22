"use strict";

var assert = require("assert");
var VarsInclude = require("../SheetRenderer/scripts/app/varsinclude.js");

function testDeepDefaults() {
    var target = {
        FOO: {
            foo: "fuga",
            nested: { b: 20 },
            list: ["local"],
            nullable: null,
            enabled: false
        }
    };
    var source = {
        FOO: {
            foo: "hoge",
            bar: "piyo",
            nested: { a: 1, b: 2 },
            list: ["included", "values"],
            nullable: "included",
            enabled: true
        }
    };

    VarsInclude.applyDeepDefaults(target, source);
    VarsInclude.stripReplaceDirectives(target);

    assert.deepStrictEqual(target, {
        FOO: {
            foo: "fuga",
            bar: "piyo",
            nested: { a: 1, b: 20 },
            list: ["local"],
            nullable: null,
            enabled: false
        }
    });
}

function testReplaceSurvivesAllIncludes() {
    var target = {
        FOO: { $replace: true, foo: "fuga" },
        EMPTY: { $replace: true }
    };

    VarsInclude.applyDeepDefaults(target, {
        FOO: { foo: "first", bar: "first" },
        EMPTY: { value: "first" }
    });
    VarsInclude.applyDeepDefaults(target, {
        FOO: { baz: "second" },
        EMPTY: { other: "second" }
    });
    VarsInclude.stripReplaceDirectives(target);

    assert.deepStrictEqual(target, {
        FOO: { foo: "fuga" },
        EMPTY: {}
    });
}

function testEarlierIncludeWins() {
    var target = {};

    VarsInclude.applyDeepDefaults(target, {
        FOO: { value: "first", firstOnly: true }
    });
    VarsInclude.applyDeepDefaults(target, {
        FOO: { value: "second", secondOnly: true }
    });

    assert.deepStrictEqual(target, {
        FOO: { value: "first", firstOnly: true, secondOnly: true }
    });
}

function testIncludedValuesAreCloned() {
    var source = { FOO: { nested: { value: 1 }, list: [1, 2] } };
    var target = {};

    VarsInclude.applyDeepDefaults(target, source);
    target.FOO.nested.value = 2;
    target.FOO.list.push(3);

    assert.deepStrictEqual(source, {
        FOO: { nested: { value: 1 }, list: [1, 2] }
    });
}

function testEntryParsing() {
    assert.deepStrictEqual(VarsInclude.parseEntry("common.yml"), {
        path: "common.yml",
        merge: "shallow"
    });
    assert.deepStrictEqual(VarsInclude.parseEntry({ path: "common.yml", merge: "deep" }), {
        path: "common.yml",
        merge: "deep"
    });
    assert.throws(function () {
        VarsInclude.parseEntry({ path: "common.yml", merge: "unknown" });
    }, /shallow.*deep/);
    assert.throws(function () {
        VarsInclude.stripReplaceDirectives({ FOO: { $replace: false } });
    }, /true/);
}

testDeepDefaults();
testReplaceSurvivesAllIncludes();
testEarlierIncludeWins();
testIncludedValuesAreCloned();
testEntryParsing();

console.log("varsinclude tests passed");
