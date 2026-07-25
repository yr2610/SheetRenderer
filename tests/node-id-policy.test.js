"use strict";

var assert = require("assert");
var NodeIdPolicy = require("../SheetRenderer/scripts/app/txt2json.js");

function testItemIdPrefix() {
    assert.deepStrictEqual(NodeIdPolicy.parseItemIdPrefix("[#] foo"), {
        auto: true,
        id: void 0,
        text: "foo"
    });
    assert.deepStrictEqual(NodeIdPolicy.parseItemIdPrefix("[#foo-id] foo"), {
        auto: false,
        id: "foo-id",
        text: "foo"
    });
    assert.deepStrictEqual(NodeIdPolicy.parseItemIdPrefix("[#]"), {
        auto: true,
        id: void 0,
        text: void 0
    });
    assert.strictEqual(NodeIdPolicy.parseItemIdPrefix("foo"), null);
}

function testGeneratedIdLine() {
    assert.strictEqual(
        NodeIdPolicy.buildItemLineWithGeneratedId("- foo", "{uid}"),
        "- [#{uid}] foo"
    );
    assert.strictEqual(
        NodeIdPolicy.buildItemLineWithGeneratedId("  * [#] foo", "abc123"),
        "  * [#abc123] foo"
    );
    assert.strictEqual(
        NodeIdPolicy.buildItemLineWithGeneratedId("+ [#old-id] foo", "new-id"),
        "+ [#new-id] foo"
    );
    assert.strictEqual(
        NodeIdPolicy.buildItemLineWithGeneratedId("foo", "abc123"),
        null
    );
}

function testIdentifiedItemRetention() {
    assert.strictEqual(NodeIdPolicy.shouldKeepIdentifiedItemWhenChildrenDisappear({
        kind: "UL",
        id: "foo-id"
    }), true);
    assert.strictEqual(NodeIdPolicy.shouldKeepIdentifiedItemWhenChildrenDisappear({
        kind: "UL"
    }), false);
    assert.strictEqual(NodeIdPolicy.shouldKeepIdentifiedItemWhenChildrenDisappear({
        kind: "H",
        id: "sheet-id"
    }), false);
}

testItemIdPrefix();
testGeneratedIdLine();
testIdentifiedItemRetention();

console.log("node ID policy tests passed");
