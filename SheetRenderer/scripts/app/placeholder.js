var Placeholder = (function () {
    var MAX_EXPANSION_DEPTH = 20;
    var RE_EMPTY = /\{\{\s*\}\}/g;
    var RE_EXPR_SOURCE = "\\{\\{\\s*([^\\}]+)\\s*\\}\\}";

    function expressionRegex() {
        return new RegExp(RE_EXPR_SOURCE, "g");
    }

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
        var regex = expressionRegex();
        var output = text.replace(regex, function (whole, raw) {
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
        var regex = expressionRegex();
        var match;
        while ((match = regex.exec(text)) !== null) {
            if (parseToken(match[1]).mode === "optionalNode") {
                return true;
            }
        }
        return false;
    }

    function findLineGuard(line, options) {
        var regex = expressionRegex();
        var match;
        var guard = null;
        var leadingLength = (line.match(/^[ \t]*/) || [""])[0].length;

        while ((match = regex.exec(line)) !== null) {
            var token = parseToken(match[1]);
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
                length: match[0].length,
                token: token
            };
        }
        return guard;
    }

    function replaceRequiredPlaceholders(line, options) {
        var drop = false;
        var regex = expressionRegex();
        var output = line.replace(regex, function (whole, raw) {
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
