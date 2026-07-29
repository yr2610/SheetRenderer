"use strict";

var assert = require("assert");
var Placeholder = require("../SheetRenderer/scripts/app/placeholder.js");

function evaluator(scope) {
    return function (expression) {
        try {
            return Function("scope", "with(scope){ return (" + expression + "); }")(scope);
        } catch (e) {
            return void 0;
        }
    };
}

function expand(text, scope, options) {
    options = options || {};
    return Placeholder.expand(text, {
        defaultParamKey: options.defaultParamKey,
        allowLineGuard: options.allowLineGuard !== false,
        evaluate: evaluator(scope || {}),
        createError: function (message) {
            return new Error(message);
        }
    });
}

function testRequiredInterpolation() {
    assert.deepStrictEqual(expand("name={{name}}, count={{count}}", {
        name: "foo",
        count: 0
    }), {
        drop: false,
        text: "name=foo, count=0"
    });

    assert.throws(function () {
        expand("{{missing}}", {});
    }, /未定義プレースホルダー.*\{\{\? missing\}\}/s);
    assert.strictEqual(expand("before {{value}} after", { value: null }).drop, true);
    assert.strictEqual(expand("before {{value}} after", { value: false }).drop, true);
    assert.deepStrictEqual(expand("before {{value}} after", { value: true }), {
        drop: false,
        text: "before  after"
    });
}

function testOptionalNodeInterpolation() {
    [false, null, void 0].forEach(function (value) {
        assert.strictEqual(expand("before {{? value}} after", { value: value }).drop, true);
    });
    assert.deepStrictEqual(expand("before {{? value}} after", { value: true }), {
        drop: false,
        text: "before  after"
    });
    assert.deepStrictEqual(expand("before {{?value}} after", { value: "x" }), {
        drop: false,
        text: "before x after"
    });
}

function testOptionalNodeRunsBeforeRequiredInterpolation() {
    assert.strictEqual(expand("{{missing}} {{? enabled}}", {
        enabled: false
    }).drop, true);
    assert.strictEqual(expand("{{?- enabled}}line {{missing}}\n{{? nodeEnabled}}other", {
        enabled: true,
        nodeEnabled: false
    }).drop, true);
}

function testOptionalLineGuard() {
    assert.deepStrictEqual(expand("first\n{{?- detail}}second\nthird", {
        detail: false
    }), {
        drop: false,
        text: "first\nthird"
    });
    assert.deepStrictEqual(expand("first\n{{?- detail}}second\nthird", {
        detail: true
    }), {
        drop: false,
        text: "first\nsecond\nthird"
    });
    assert.strictEqual(expand("{{?- detail}}only", { detail: false }).drop, true);
    assert.strictEqual(expand("{{?- detail}}only", {}).drop, true);
    assert.strictEqual(expand("{{?- detail}}only", { detail: null }).drop, true);
    assert.throws(function () {
        expand("{{?- detail}}only", { detail: "yes" });
    }, /boolean のみ/);
}

function testDeletedLineIsNotEvaluated() {
    assert.deepStrictEqual(expand("first\n{{?- detail}}value={{missing}}\nthird", {
        detail: false
    }), {
        drop: false,
        text: "first\nthird"
    });
}

function testLineGuardPlacement() {
    assert.throws(function () {
        expand("prefix {{?- detail}}suffix", { detail: false });
    }, /行頭/);
    assert.throws(function () {
        expand("{{?- first}}{{?- second}}suffix", { first: true, second: true });
    }, /行頭|複数/);
    assert.throws(function () {
        expand("{{?- detail}}suffix", { detail: true }, { allowLineGuard: false });
    }, /ノードのテキストでのみ/);
}

function testNegativeOptionalExpressionCompatibility() {
    assert.deepStrictEqual(expand("{{?-value}}", { value: 2 }), {
        drop: false,
        text: "-2"
    });
    assert.deepStrictEqual(expand("{{? -value}}", { value: 2 }), {
        drop: false,
        text: "-2"
    });
}

function testTemplateDefaultPlaceholder() {
    assert.deepStrictEqual(expand("value={{}}", { $value: "x" }, {
        defaultParamKey: "$value"
    }), {
        drop: false,
        text: "value=x"
    });
}

function testJavaScriptWithNestedClosingBraces() {
    assert.deepStrictEqual(expand("{{({ outer: { value: value } }).outer.value}}", {
        value: "nested"
    }), {
        drop: false,
        text: "nested"
    });
    assert.deepStrictEqual(expand("{{/}}/.test(value)}}", {
        value: "a}}b"
    }), {
        drop: false,
        text: ""
    });
    assert.deepStrictEqual(expand("{{'left}}right'}}"), {
        drop: false,
        text: "left}}right"
    });
}

function testTemplateLiteralInterpolation() {
    assert.deepStrictEqual(expand("{{`foo${bar}`}}", { bar: "BAR" }), {
        drop: false,
        text: "fooBAR"
    });
    assert.deepStrictEqual(expand("{{`outer ${enabled ? `inner ${value}` : 'off'}`}}", {
        enabled: true,
        value: 3
    }), {
        drop: false,
        text: "outer inner 3"
    });
    assert.deepStrictEqual(expand("first={{`a${value}`}}, second={{value + 1}}", {
        value: 2
    }), {
        drop: false,
        text: "first=a2, second=3"
    });
}

testRequiredInterpolation();
testOptionalNodeInterpolation();
testOptionalNodeRunsBeforeRequiredInterpolation();
testOptionalLineGuard();
testDeletedLineIsNotEvaluated();
testLineGuardPlacement();
testNegativeOptionalExpressionCompatibility();
testTemplateDefaultPlaceholder();
testJavaScriptWithNestedClosingBraces();
testTemplateLiteralInterpolation();

console.log("placeholder tests passed");
