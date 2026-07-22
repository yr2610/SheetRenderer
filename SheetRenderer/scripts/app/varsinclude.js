var VarsInclude = (function () {
    var hasOwn = Object.prototype.hasOwnProperty;
    var objectToString = Object.prototype.toString;

    function isMapping(value) {
        return value !== null
            && typeof value === "object"
            && !Array.isArray(value)
            && objectToString.call(value) === "[object Object]";
    }

    function setOwnProperty(target, key, value) {
        if (key === "__proto__") {
            Object.defineProperty(target, key, {
                configurable: true,
                enumerable: true,
                writable: true,
                value: value
            });
            return;
        }

        target[key] = value;
    }

    function cloneValue(value) {
        if (Array.isArray(value)) {
            return value.map(cloneValue);
        }

        if (isMapping(value)) {
            var clone = {};
            Object.keys(value).forEach(function (key) {
                setOwnProperty(clone, key, cloneValue(value[key]));
            });
            return clone;
        }

        if (objectToString.call(value) === "[object Date]") {
            return new Date(value.getTime());
        }

        return value;
    }

    function isReplaceMapping(value) {
        if (!isMapping(value) || !hasOwn.call(value, "$replace")) {
            return false;
        }

        if (value.$replace !== true) {
            throw new Error("vars.yml の $replace には true を指定してください。");
        }

        return true;
    }

    // target に存在しない値だけを source から取り込む。
    // mapping 同士だけを再帰処理し、配列とスカラーはひとつの値として扱う。
    function applyDeepDefaults(target, source) {
        if (!isMapping(target) || !isMapping(source) || isReplaceMapping(target)) {
            return target;
        }

        Object.keys(source).forEach(function (key) {
            if (key === "$replace") {
                return;
            }

            if (!hasOwn.call(target, key)) {
                setOwnProperty(target, key, cloneValue(source[key]));
                return;
            }

            if (isMapping(target[key]) && isMapping(source[key])) {
                applyDeepDefaults(target[key], source[key]);
            }
        });

        return target;
    }

    // $replace はマージ制御専用なので、最終的な変数データから取り除く。
    function stripReplaceDirectives(value) {
        if (Array.isArray(value)) {
            value.forEach(stripReplaceDirectives);
            return value;
        }

        if (!isMapping(value)) {
            return value;
        }

        if (hasOwn.call(value, "$replace")) {
            if (value.$replace !== true) {
                throw new Error("vars.yml の $replace には true を指定してください。");
            }
            delete value.$replace;
        }

        Object.keys(value).forEach(function (key) {
            stripReplaceDirectives(value[key]);
        });

        return value;
    }

    function parseEntry(entry) {
        if (typeof entry === "string") {
            return {
                path: entry,
                merge: "shallow"
            };
        }

        if (!isMapping(entry) || typeof entry.path !== "string" || !entry.path.trim()) {
            throw new Error("vars.yml の $include はパス文字列、または path を持つオブジェクトで指定してください。");
        }

        var merge = typeof entry.merge === "undefined" ? "shallow" : entry.merge;
        if (merge !== "shallow" && merge !== "deep") {
            throw new Error("vars.yml の $include.merge には shallow または deep を指定してください。");
        }

        return {
            path: entry.path,
            merge: merge
        };
    }

    return {
        applyDeepDefaults: applyDeepDefaults,
        parseEntry: parseEntry,
        stripReplaceDirectives: stripReplaceDirectives
    };
}());

if (typeof module !== "undefined" && module.exports) {
    module.exports = VarsInclude;
}
