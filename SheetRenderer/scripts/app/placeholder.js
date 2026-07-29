var Placeholder = (function () {
    var MAX_EXPANSION_DEPTH = 20;
    var RE_EMPTY = /\{\{\s*\}\}/g;
    var expressionSyntaxCache = Object.create(null);

    function parseToken(raw) {
        var text = String(raw || "").trim();
        var lineMatch = text.match(/^\?-\s+([\s\S]+)$/);
        if (lineMatch) {
            return {
                mode: "optionalLine",
                expression: lineMatch[1].trim(),
                source: text
            };
        }
        if (text.charAt(0) === "?") {
            return {
                mode: "optionalNode",
                expression: text.slice(1).trim(),
                source: text
            };
        }
        return {
            mode: "required",
            expression: text,
            source: text
        };
    }

    function hasValidExpressionSyntax(raw) {
        var token = parseToken(raw);
        var expression = token.expression;
        if (!expression) {
            return false;
        }
        if (Object.prototype.hasOwnProperty.call(expressionSyntaxCache, expression)) {
            return expressionSyntaxCache[expression];
        }

        try {
            // evalExprWithScope と同じ形でコンパイルだけ行う。式そのものは実行しない。
            new Function("scope", "with(scope){ return (" + expression + "); }");
            expressionSyntaxCache[expression] = true;
        } catch (e) {
            expressionSyntaxCache[expression] = false;
        }
        return expressionSyntaxCache[expression];
    }

    function findPlaceholderAt(text, start) {
        var close = text.indexOf("}}", start + 2);
        while (close !== -1) {
            var raw = text.slice(start + 2, close);
            if (hasValidExpressionSyntax(raw)) {
                return {
                    index: start,
                    length: close + 2 - start,
                    whole: text.slice(start, close + 2),
                    raw: raw
                };
            }
            // "}}}" も調べられるよう、候補は1文字ずつ進める。
            close = text.indexOf("}}", close + 1);
        }
        return null;
    }

    function findPlaceholders(text) {
        var matches = [];
        var searchFrom = 0;
        var start;

        while ((start = text.indexOf("{{", searchFrom)) !== -1) {
            var match = findPlaceholderAt(text, start);
            if (match !== null) {
                matches.push(match);
                searchFrom = match.index + match.length;
            } else {
                searchFrom = start + 2;
            }
        }
        return matches;
    }

    function replacePlaceholders(text, replacer) {
        var matches = findPlaceholders(text);
        if (matches.length === 0) {
            return text;
        }

        var parts = [];
        var copiedUntil = 0;
        for (var i = 0; i < matches.length; i++) {
            var match = matches[i];
            parts.push(text.slice(copiedUntil, match.index));
            parts.push(replacer(match.whole, match.raw, match.index));
            copiedUntil = match.index + match.length;
        }
        parts.push(text.slice(copiedUntil));
        return parts.join("");
    }

    function fail(options, message) {
        if (options && typeof options.createError === "function") {
            throw options.createError(message);
        }
        throw new Error(message);
    }

    function evaluate(token, options) {
        if (!token.expression) {
            fail(options, "プレースホルダーの式がありません: {{" + token.source + "}}");
        }
        if (token.expression.charAt(token.expression.length - 1) === "!") {
            fail(options, "プレースホルダー '{{" + token.source + "}}' の '!' 指定は廃止されました。");
        }
        return options.evaluate(token.expression);
    }

    function requiredResult(token, value, options) {
        if (value === void 0) {
            fail(
                options,
                "未定義プレースホルダー: " + token.expression +
                "\n未定義を意図的に許可する場合は {{? " + token.expression + "}} を使用してください。"
            );
        }
        // 移行期間中は従来互換を維持する。
        // false/null はノード削除、true はガードとして空文字にする。
        if (value === false || value === null) {
            return { drop: true, text: "" };
        }
        return {
            drop: false,
            text: value === true ? "" : String(value)
        };
    }

    function replaceOptionalNodes(text, options) {
        var drop = false;
        var changed = false;
        var output = replacePlaceholders(text, function (whole, raw) {
            var token = parseToken(raw);
            var value;
            if (token.mode !== "optionalNode") {
                return whole;
            }
            if (drop) {
                return "";
            }

            value = evaluate(token, options);
            changed = true;
            if (value === false || value === null || value === void 0) {
                drop = true;
                return "";
            }
            return value === true ? "" : String(value);
        });

        return {
            drop: drop,
            changed: changed,
            text: output
        };
    }

    function containsOptionalNode(text) {
        var matches = findPlaceholders(text);
        for (var i = 0; i < matches.length; i++) {
            if (parseToken(matches[i].raw).mode === "optionalNode") {
                return true;
            }
        }
        return false;
    }

    function findLineGuard(line, options) {
        var matches = findPlaceholders(line);
        var guard = null;
        var leadingLength = (line.match(/^[ \t]*/) || [""])[0].length;

        for (var i = 0; i < matches.length; i++) {
            var match = matches[i];
            var token = parseToken(match.raw);
            if (token.mode !== "optionalLine") {
                continue;
            }
            if (!options.allowLineGuard) {
                fail(options, "行ガード '{{" + token.source + "}}' はノードのテキストでのみ使用できます。");
            }
            if (match.index !== leadingLength) {
                fail(options, "行ガード '{{" + token.source + "}}' は行頭に記述してください。");
            }
            if (guard !== null) {
                fail(options, "1行に複数の行ガードは記述できません。");
            }
            guard = {
                index: match.index,
                length: match.length,
                token: token
            };
        }
        return guard;
    }

    function replaceRequiredPlaceholders(line, options) {
        var drop = false;
        var output = replacePlaceholders(line, function (whole, raw) {
            var token = parseToken(raw);
            var result;
            if (token.mode !== "required") {
                return whole;
            }
            if (drop) {
                return "";
            }
            result = requiredResult(token, evaluate(token, options), options);
            if (result.drop) {
                drop = true;
                return "";
            }
            return result.text;
        });
        return { drop: drop, text: output };
    }

    function replaceLines(text, options) {
        var lines = String(text).split(/\r?\n/);
        var outputLines = [];

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            var guard = findLineGuard(line, options);
            if (guard !== null) {
                var value = evaluate(guard.token, options);
                if (value === false || value === null || value === void 0) {
                    continue;
                }
                if (value !== true) {
                    fail(
                        options,
                        "行ガード '{{" + guard.token.source + "}}' は boolean のみ使用できます。"
                    );
                }
                line = line.slice(0, guard.index) + line.slice(guard.index + guard.length);
            }
            var requiredResultForLine = replaceRequiredPlaceholders(line, options);
            if (requiredResultForLine.drop) {
                return { drop: true, text: void 0 };
            }
            outputLines.push(requiredResultForLine.text);
        }

        return {
            drop: outputLines.length === 0,
            text: outputLines.join("\n")
        };
    }

    function expand(text, options) {
        options = options || {};
        if (typeof options.evaluate !== "function") {
            throw new Error("Placeholder.expand には evaluate が必要です。");
        }
        if (text === void 0 || text === null) {
            return { drop: false, text: text };
        }

        var current = String(text);
        for (var depth = 0; depth < MAX_EXPANSION_DEPTH; depth++) {
            var before = current;
            if (options.defaultParamKey) {
                current = current.replace(RE_EMPTY, "{{" + options.defaultParamKey + "}}");
            }

            var optionalResult = replaceOptionalNodes(current, options);
            if (optionalResult.drop) {
                return { drop: true, text: void 0 };
            }
            current = optionalResult.text;

            // 置換結果に新しいノード省略プレースホルダーが現れた場合は、
            // 必須プレースホルダーより先に評価するため次のパスへ送る。
            if (current !== before && containsOptionalNode(current)) {
                continue;
            }

            var lineResult = replaceLines(current, options);
            if (lineResult.drop) {
                return { drop: true, text: void 0 };
            }
            current = lineResult.text;

            if (current === before) {
                return { drop: false, text: current };
            }
        }

        fail(
            options,
            "プレースホルダー展開が " + MAX_EXPANSION_DEPTH +
            " 回を超えました。循環参照の可能性があります。"
        );
    }

    return {
        expand: expand,
        parseToken: parseToken
    };
}());

if (typeof module !== "undefined" && module.exports) {
    module.exports = Placeholder;
}
