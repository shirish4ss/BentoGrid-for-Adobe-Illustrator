/**
 * ================================================================
 *  ██████╗ ███████╗███╗   ██╗████████╗ ██████╗  ██████╗ ██████╗ ██╗██████╗
 *  ██╔══██╗██╔════╝████╗  ██║╚══██╔══╝██╔═══██╗██╔════╝ ██╔══██╗██║██╔══██╗
 *  ██████╔╝█████╗  ██╔██╗ ██║   ██║   ██║   ██║██║  ███╗██████╔╝██║██║  ██║
 *  ██╔══██╗██╔══╝  ██║╚██╗██║   ██║   ██║   ██║██║   ██║██╔══██╗██║██║  ██║
 *  ██████╔╝███████╗██║ ╚████║   ██║   ╚██████╔╝╚██████╔╝██║  ██║██║██████╔╝
 *  ╚═════╝ ╚══════╝╚═╝  ╚═══╝   ╚═╝    ╚═════╝  ╚═════╝ ╚═╝  ╚═╝╚═╝╚═════╝
 *                         by  S H I R I S H   S A L V E
 * ================================================================
 *  Version  : 3.0
 *  Author   : Shirish Salve
 *  Studio   : 2Advanced Studios
 *  Year     : 2026
 * ----------------------------------------------------------------
 *  v3.0 New Features:
 *   - Brand new BENTOGRID logo header (text-only, no image dependency)
 *   - 50+ bento patterns across 2–6 columns
 *   - Color Palette Presets: Dark, Light, Vibrant, Pastel, Monochrome
 *   - Random Pattern Generator (Shuffle button)
 *   - Live Grid Preview Panel inside dialog
 *   - Save / Load Settings (JSON to temp file)
 *   - Improved tabbed UI with 3 tabs
 *   - Cleaner codebase with proper error handling
 * ================================================================
 */

(function () {
    'use strict';

    // ── Guard ────────────────────────────────────────────────────
    if (app.documents.length === 0) {
        alert('Please open a document first.');
        return;
    }
    var doc = app.activeDocument;

    // ================================================================
    //  COLOR UTILITIES
    // ================================================================
    function clamp01(v) { return Math.max(0, Math.min(1, v)); }

    function createColor(rgb) {
        var c = new RGBColor();
        c.red   = Math.round(clamp01(rgb[0]) * 255);
        c.green = Math.round(clamp01(rgb[1]) * 255);
        c.blue  = Math.round(clamp01(rgb[2]) * 255);
        return c;
    }

    function colorToRGB01(c) {
        try {
            if (!c) return null;
            if (c.typename === 'RGBColor')  return [c.red/255, c.green/255, c.blue/255];
            if (c.typename === 'GrayColor') { var v=1-(c.gray/100); return [v,v,v]; }
            if (c.typename === 'CMYKColor') {
                var C=c.cyan/100, M=c.magenta/100, Y=c.yellow/100, K=c.black/100;
                return [(1-C)*(1-K), (1-M)*(1-K), (1-Y)*(1-K)];
            }
            if (c.typename === 'SpotColor') return colorToRGB01(c.spot.color);
        } catch(e) {}
        return null;
    }

    function rgbToHex(rgb) {
        function h(v) { var s=Math.round(clamp01(v)*255).toString(16); return s.length<2?'0'+s:s; }
        return '#'+h(rgb[0])+h(rgb[1])+h(rgb[2]);
    }

    function hexToRGB01(hex) {
        hex = hex.replace('#','');
        if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
        if (hex.length !== 6) return null;
        return [parseInt(hex.substr(0,2),16)/255, parseInt(hex.substr(2,2),16)/255, parseInt(hex.substr(4,2),16)/255];
    }

    function getIllustratorForegroundRGB() {
        try {
            if (doc.defaultFillColor) {
                var rgb = colorToRGB01(doc.defaultFillColor);
                if (rgb) return rgb;
            }
        } catch(e) {}
        return null;
    }

    // ================================================================
    //  COLOR PALETTE PRESETS
    // ================================================================
    var colorPalettes = {
        'Custom RGB':     null,
        'Deep Dark':      [[0.08,0.08,0.10], [0.14,0.14,0.18], [0.20,0.20,0.26]],
        'Charcoal Smoke': [[0.18,0.18,0.20], [0.28,0.28,0.30], [0.38,0.38,0.40]],
        'Pure Black':     [[0.00,0.00,0.00], [0.10,0.10,0.10], [0.20,0.20,0.20]],
        'Clean White':    [[1.00,1.00,1.00], [0.95,0.95,0.95], [0.88,0.88,0.88]],
        'Soft Ivory':     [[0.98,0.97,0.93], [0.93,0.91,0.86], [0.86,0.84,0.78]],
        'Neon Pulse':     [[0.00,1.00,0.70], [0.08,0.08,0.12], [1.00,0.20,0.60]],
        'Ocean Blue':     [[0.04,0.35,0.78], [0.10,0.55,0.92], [0.02,0.18,0.45]],
        'Sunset Orange':  [[0.98,0.42,0.08], [0.96,0.68,0.20], [0.80,0.20,0.10]],
        'Forest Green':   [[0.08,0.45,0.22], [0.14,0.62,0.35], [0.04,0.28,0.14]],
        'Rose Blush':     [[0.95,0.60,0.68], [0.98,0.82,0.86], [0.82,0.38,0.50]],
        'Lavender Mist':  [[0.72,0.65,0.92], [0.88,0.84,0.98], [0.52,0.44,0.78]],
        'Sky Pastel':     [[0.72,0.88,0.98], [0.90,0.96,1.00], [0.50,0.74,0.92]],
        'Mint Fresh':     [[0.60,0.94,0.80], [0.84,0.98,0.92], [0.30,0.78,0.62]],
        'Warm Sand':      [[0.96,0.88,0.72], [0.98,0.93,0.84], [0.82,0.72,0.52]],
        'Mono Grey':      [[0.50,0.50,0.50], [0.70,0.70,0.70], [0.30,0.30,0.30]],
        'Mono Slate':     [[0.35,0.40,0.46], [0.55,0.60,0.66], [0.18,0.22,0.28]],
        'Purple Haze':    [[0.45,0.15,0.75], [0.68,0.40,0.92], [0.25,0.05,0.50]],
        'Coral Reef':     [[0.98,0.50,0.40], [0.98,0.72,0.62], [0.82,0.30,0.22]],
        'Deep Teal':      [[0.04,0.50,0.52], [0.08,0.72,0.74], [0.02,0.30,0.32]]
    };

    var paletteNames = [];
    for (var pk in colorPalettes) if (colorPalettes.hasOwnProperty(pk)) paletteNames.push(pk);

    // ================================================================
    //  BENTO PATTERNS  (50+)
    // ================================================================
    var bentoPatterns = {

        // ── 1 Column ───────────────────────────────────────────────
        'Single Full (1×1)': [
            [ {cs:1,rs:1} ]
        ],
        'Single Stack (1×2)': [
            [ {cs:1,rs:1} ],
            [ {cs:1,rs:1} ]
        ],
        'Single Stack (1×3)': [
            [ {cs:1,rs:1} ],
            [ {cs:1,rs:1} ],
            [ {cs:1,rs:1} ]
        ],

        // ── 2 Column ───────────────────────────────────────────────
        'Split Equal (2×1)': [
            [ {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Split Equal (2×2)': [
            [ {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Split Equal (2×3)': [
            [ {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Tall Left (2×2)': [
            [ {cs:1,rs:2}, {cs:1,rs:1} ],
            [               {cs:1,rs:1} ]
        ],
        'Tall Right (2×2)': [
            [ {cs:1,rs:1}, {cs:1,rs:2} ],
            [ {cs:1,rs:1}               ]
        ],
        'Wide Top (2×2)': [
            [ {cs:2,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Wide Bottom (2×2)': [
            [ {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:2,rs:1} ]
        ],

        // ── 3 Column ───────────────────────────────────────────────
        'Hero Left (3×2)': [
            [ {cs:2,rs:2}, {cs:1,rs:1} ],
            [               {cs:1,rs:1} ]
        ],
        'Hero Right (3×2)': [
            [ {cs:1,rs:1}, {cs:2,rs:2} ],
            [ {cs:1,rs:1}               ]
        ],
        'Hero Top (3×2)': [
            [ {cs:3,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Hero Bottom (3×2)': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:3,rs:1} ]
        ],
        'Classic 3×2': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Classic 3×3': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Center Hero (3×3)': [
            [ {cs:1,rs:1}, {cs:1,rs:2}, {cs:1,rs:1} ],
            [ {cs:1,rs:1},               {cs:1,rs:1} ],
            [ {cs:3,rs:1} ]
        ],
        'T-Top (3×3)': [
            [ {cs:3,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:2,rs:1} ]
        ],
        'Feature Left Stack (3×3)': [
            [ {cs:2,rs:2}, {cs:1,rs:1} ],
            [               {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Pinwheel (3×3)': [
            [ {cs:2,rs:1}, {cs:1,rs:2} ],
            [ {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:2,rs:1} ]
        ],
        'Mosaic (3×3)': [
            [ {cs:1,rs:2}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [               {cs:2,rs:1}               ],
            [ {cs:2,rs:1},               {cs:1,rs:1} ]
        ],
        'Brick Left (3×3)': [
            [ {cs:1,rs:1}, {cs:2,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:2,rs:1},              {cs:1,rs:1} ]
        ],
        'Tall Center (3×3)': [
            [ {cs:1,rs:1}, {cs:1,rs:3}, {cs:1,rs:1} ],
            [ {cs:1,rs:1},               {cs:1,rs:1} ],
            [ {cs:1,rs:1},               {cs:1,rs:1} ]
        ],

        // ── 4 Column ───────────────────────────────────────────────
        'Classic 4×2': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Classic 4×3': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Hero Top Wide (4×2)': [
            [ {cs:4,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Hero Bottom Wide (4×2)': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:4,rs:1} ]
        ],
        'Feature Left 4×2': [
            [ {cs:2,rs:2}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [               {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Feature Right 4×2': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:2,rs:2} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}               ]
        ],
        'Feature Center 4×2': [
            [ {cs:1,rs:1}, {cs:2,rs:2}, {cs:1,rs:1} ],
            [ {cs:1,rs:1},               {cs:1,rs:1} ]
        ],
        'Wing Span (4×3)': [
            [ {cs:1,rs:2}, {cs:2,rs:1}, {cs:1,rs:2} ],
            [               {cs:2,rs:1}               ],
            [ {cs:1,rs:1}, {cs:2,rs:1}, {cs:1,rs:1} ]
        ],
        'Dashboard (4×3)': [
            [ {cs:2,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:2,rs:1} ],
            [ {cs:1,rs:1}, {cs:2,rs:1}, {cs:1,rs:1} ]
        ],
        'Tall Feature (4×3)': [
            [ {cs:1,rs:3}, {cs:1,rs:1}, {cs:2,rs:1} ],
            [               {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [               {cs:3,rs:1}               ]
        ],
        'Landscape Hero (4×2)': [
            [ {cs:1,rs:1}, {cs:2,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Stagger Rows (4×3)': [
            [ {cs:1,rs:1}, {cs:2,rs:1}, {cs:1,rs:1} ],
            [ {cs:2,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:2,rs:1} ]
        ],
        'Big + Small (4×3)': [
            [ {cs:3,rs:2}, {cs:1,rs:1} ],
            [               {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],

        // ── 5 Column ───────────────────────────────────────────────
        'Classic 5×2': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Classic 5×3': [
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Panorama 5×2': [
            [ {cs:5,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Feature Left 5×2': [
            [ {cs:2,rs:2}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [               {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Feature Center 5×2': [
            [ {cs:1,rs:1}, {cs:3,rs:2}, {cs:1,rs:1} ],
            [ {cs:1,rs:1},               {cs:1,rs:1} ]
        ],
        'Wide Feature Center 5×3': [
            [ {cs:1,rs:1}, {cs:3,rs:2}, {cs:1,rs:1} ],
            [ {cs:1,rs:1},               {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Alt Stagger (5×3)': [
            [ {cs:2,rs:1}, {cs:1,rs:1}, {cs:2,rs:1} ],
            [ {cs:1,rs:1}, {cs:3,rs:1}, {cs:1,rs:1} ],
            [ {cs:2,rs:1}, {cs:1,rs:1}, {cs:2,rs:1} ]
        ],

        // ── 6 Column ───────────────────────────────────────────────
        'Classic 6×2': [
            [ {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ],
            [ {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ]
        ],
        'Classic 6×3': [
            [ {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ],
            [ {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ],
            [ {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ]
        ],
        'Feature Left 6×2': [
            [ {cs:3,rs:2},{cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ],
            [              {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ]
        ],
        'Full Bleed 6×3': [
            [ {cs:6,rs:1} ],
            [ {cs:3,rs:1},{cs:2,rs:1},{cs:1,rs:1} ],
            [ {cs:1,rs:1},{cs:2,rs:1},{cs:3,rs:1} ]
        ],
        'Hexagonal Flow (6×3)': [
            [ {cs:2,rs:1},{cs:2,rs:1},{cs:2,rs:1} ],
            [ {cs:1,rs:1},{cs:2,rs:1},{cs:2,rs:1},{cs:1,rs:1} ],
            [ {cs:2,rs:1},{cs:2,rs:1},{cs:2,rs:1} ]
        ],
        'Editorial Mix (6×3)': [
            [ {cs:4,rs:2},{cs:2,rs:1} ],
            [              {cs:1,rs:1},{cs:1,rs:1} ],
            [ {cs:2,rs:1},{cs:2,rs:1},{cs:2,rs:1} ]
        ],

        // ── Magazine / Editorial ────────────────────────────────────
        'Editorial 2+1': [
            [ {cs:2,rs:1}, {cs:1,rs:2} ],
            [ {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Full Bleed Strip': [
            [ {cs:4,rs:1} ],
            [ {cs:2,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:2,rs:1} ]
        ],
        'Magazine Cover': [
            [ {cs:3,rs:2} ],
            [              ],
            [ {cs:1,rs:1}, {cs:1,rs:1}, {cs:1,rs:1} ]
        ],
        'Asymmetric Split': [
            [ {cs:3,rs:1}, {cs:1,rs:3} ],
            [ {cs:1,rs:1}, {cs:2,rs:1} ],
            [ {cs:2,rs:1}, {cs:1,rs:1} ]
        ],
        'Power T': [
            [ {cs:4,rs:1} ],
            [ {cs:1,rs:2}, {cs:2,rs:1}, {cs:1,rs:2} ],
            [               {cs:2,rs:1}               ]
        ],
        'Spotlight Left': [
            [ {cs:1,rs:4}, {cs:3,rs:2} ],
            [               {cs:1,rs:1},{cs:1,rs:1},{cs:1,rs:1} ],
            [               {cs:3,rs:1} ]
        ]
    };

    var patternNames = [];
    for (var pk2 in bentoPatterns) if (bentoPatterns.hasOwnProperty(pk2)) patternNames.push(pk2);

    // ================================================================
    //  SETTINGS  (load / save to temp JSON)
    // ================================================================
    var SETTINGS_FILE_PATH = Folder.temp.fsName + '/bentogrid_shirish_settings.json';

    var settings = {
        gridWidth:          1200,
        gridHeight:         600,
        gutter:             24,
        useRoundedCorners:  false,
        cornerRadius:       16,
        fillColor:          [0.12, 0.12, 0.12],
        paletteIndex:       0,
        strokeColor:        [1, 1, 1],
        strokeWidth:        1,
        useAutoFill:        true,
        useStroke:          false,
        createBackground:   false,
        backgroundPadding:  24,
        bgColor:            [0.95, 0.95, 0.95],
        selectedPattern:    patternNames[0],
        customPatternJSON:  '',
        simplePattern:      '',
        namingPrefix:       'Cell',
        groupName:          'Bento Grid'
    };

    // ── ExtendScript-safe JSON serializer (no native JSON object) ────
    function esStringify(val) {
        if (val === null || val === undefined) return 'null';
        var t = typeof val;
        if (t === 'boolean') return val ? 'true' : 'false';
        if (t === 'number')  return isFinite(val) ? String(val) : 'null';
        if (t === 'string') {
            return '"' + val
                .replace(/\\/g, '\\\\')
                .replace(/"/g,  '\\"')
                .replace(/\n/g, '\\n')
                .replace(/\r/g, '\\r')
                .replace(/\t/g, '\\t') + '"';
        }
        if (val instanceof Array) {
            var aout = [];
            for (var ai = 0; ai < val.length; ai++) aout.push(esStringify(val[ai]));
            return '[' + aout.join(',') + ']';
        }
        if (t === 'object') {
            var oout = [];
            for (var ok in val) {
                if (val.hasOwnProperty(ok)) {
                    oout.push(esStringify(String(ok)) + ':' + esStringify(val[ok]));
                }
            }
            return '{' + oout.join(',') + '}';
        }
        return 'null';
    }

    function saveSettings() {
        try {
            var f = File(SETTINGS_FILE_PATH);
            f.encoding = 'UTF-8';
            f.open('w');
            f.write(esStringify(settings));
            f.close();
        } catch(e) { alert('Could not save settings: ' + e); }
    }

    function loadSettings() {
        try {
            var f = File(SETTINGS_FILE_PATH);
            if (!f.exists) return;
            f.encoding = 'UTF-8';
            f.open('r');
            var raw = f.read();
            f.close();
            var loaded = eval('(' + raw + ')');
            for (var key in loaded) {
                if (loaded.hasOwnProperty(key) && settings.hasOwnProperty(key)) {
                    settings[key] = loaded[key];
                }
            }
        } catch(e) {}
    }

    // ================================================================
    //  PATTERN UTILITIES
    // ================================================================
    function parseSimplePattern(str) {
        var rows = str.split('|'), out = [];
        for (var r = 0; r < rows.length; r++) {
            var cols = rows[r].split(','), prow = [];
            for (var i = 0; i < cols.length; i++) {
                var v = parseInt(cols[i], 10);
                if (!isNaN(v) && v > 0) prow.push({cs: v, rs: 1});
            }
            if (prow.length) out.push(prow);
        }
        return out;
    }

    function parseJSONPattern(str) {
        try {
            var p = eval('(' + str + ')');
            if (!(p instanceof Array)) throw 'Pattern must be an array of rows.';
            return p;
        } catch(e) { alert('Custom pattern parse error: ' + e); return null; }
    }

    function randomPattern() {
        var cols   = Math.floor(Math.random() * 4) + 2;   // 2–5 cols
        var rows   = Math.floor(Math.random() * 3) + 2;   // 2–4 rows
        var occ    = {};
        var out    = [];
        function key(r,c){ return r+'_'+c; }
        for (var r = 0; r < rows; r++) {
            var row = [], c = 0;
            while (c < cols) {
                if (occ[key(r,c)]) { c++; continue; }
                var maxCS = 1;
                while (maxCS < cols - c && !occ[key(r, c + maxCS)]) maxCS++;
                var cs = (maxCS > 1 && Math.random() > 0.5) ? (Math.floor(Math.random() * maxCS) + 1) : 1;
                var maxRS = 1;
                if (r < rows - 1) maxRS = (Math.random() > 0.6) ? 2 : 1;
                var rs = maxRS;
                row.push({cs: cs, rs: rs});
                for (var dr = 0; dr < rs; dr++) {
                    for (var dc = 0; dc < cs; dc++) {
                        occ[key(r + dr, c + dc)] = true;
                    }
                }
                c += cs;
            }
            if (row.length) out.push(row);
        }
        return out;
    }

    // ================================================================
    //  DIALOG
    // ================================================================
    function createDialog() {
        loadSettings();

        // Sync artboard size
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        settings.gridWidth  = ab[2] - ab[0];
        settings.gridHeight = ab[1] - ab[3];

        var dlg = new Window('dialog', 'BentoGrid  \u2014  by Shirish Salve  v3.0');
        dlg.orientation    = 'column';
        dlg.alignChildren  = 'fill';
        dlg.spacing        = 4;

        // ── LOGO HEADER ──────────────────────────────────────────────
        var logoGrp = dlg.add('group');
        logoGrp.orientation   = 'column';
        logoGrp.alignChildren = 'center';
        logoGrp.alignment     = 'fill';
        logoGrp.margins       = [0, 4, 0, 2];
        logoGrp.spacing       = 2;

        var logoTitle = logoGrp.add('statictext', undefined, 'BentoGrid  by Shirish Salve');
        try { logoTitle.graphics.font = ScriptUI.newFont('Arial', ScriptUI.FontStyle.BOLD, 14); } catch(_) {}

        var logoSub = logoGrp.add('statictext', undefined, 'v3.0  \u2014  Adobe Illustrator Script  \u00A9 2026');
        try { logoSub.graphics.font = ScriptUI.newFont('Arial', ScriptUI.FontStyle.ITALIC, 10); } catch(_) {}

        // ── TABS ─────────────────────────────────────────────────────
        var tabs = dlg.add('tabbedpanel');
        tabs.alignment        = 'fill';
        tabs.alignChildren    = 'fill';
        tabs.preferredSize    = [540, 490];
        tabs.maximumSize      = [540, 490];

        // ════════════════════════════════════════════════════════════
        //  TAB 1 — GRID & PATTERN
        // ════════════════════════════════════════════════════════════
        var tab1 = tabs.add('tab', undefined, ' \u25A6 Grid & Pattern ');
        tab1.orientation    = 'column';
        tab1.alignChildren  = 'fill';
        tab1.margins        = [8, 10, 8, 4];
        tab1.spacing        = 5;

        // Grid size
        var dimPanel = tab1.add('panel', undefined, 'Grid Size');
        dimPanel.orientation   = 'column';
        dimPanel.alignChildren = 'fill';
        dimPanel.margins       = [8, 12, 8, 6];

        var dimRow = dimPanel.add('group');
        dimRow.add('statictext', undefined, 'Width:');
        var widthInput = dimRow.add('edittext', undefined, String(Math.round(settings.gridWidth)));
        widthInput.characters = 8;
        dimRow.add('statictext', undefined, '\u00D7   Height:');
        var heightInput = dimRow.add('edittext', undefined, String(Math.round(settings.gridHeight)));
        heightInput.characters = 8;
        dimRow.add('statictext', undefined, 'px');

        var dimHint = dimPanel.add('statictext', undefined, 'Auto-synced from active artboard on open.');
        try { dimHint.graphics.font = ScriptUI.newFont(dimHint.graphics.font.name, ScriptUI.FontStyle.ITALIC, 10); } catch(_) {}

        // Gutter
        var gutterPanel = tab1.add('panel', undefined, 'Gutter');
        gutterPanel.orientation   = 'column';
        gutterPanel.alignChildren = 'fill';
        gutterPanel.margins       = [8, 12, 8, 6];

        var gutterRow = gutterPanel.add('group');
        gutterRow.add('statictext', undefined, 'Gutter:');
        var gutterInput = gutterRow.add('edittext', undefined, String(settings.gutter));
        gutterInput.characters = 6;
        gutterRow.add('statictext', undefined, 'px');

        var gpRow = gutterPanel.add('group');
        gpRow.add('statictext', undefined, 'Quick preset:');
        var gpDrop = gpRow.add('dropdownlist', undefined, ['Minimal (8px)', 'Tight (12px)', 'Comfort (24px)', 'Loose (48px)', 'Spacious (64px)', 'None (0px)', 'Custom']);
        gpDrop.preferredSize.width = 200;

        function getPresetPx(idx) {
            return [8,12,24,48,64,0,null][idx];
        }
        function getPresetIdx(px) {
            switch(px){case 8:return 0;case 12:return 1;case 24:return 2;case 48:return 3;case 64:return 4;case 0:return 5;default:return 6;}
        }
        gpDrop.selection = getPresetIdx(settings.gutter);
        gpDrop.onChange = function() {
            var px = getPresetPx(gpDrop.selection.index);
            if (px !== null) { gutterInput.text = String(px); }
            updatePreview();
        };
        gutterInput.onChange = function() { gpDrop.selection = getPresetIdx(parseInt(gutterInput.text,10)); updatePreview(); };

        // Pattern selector
        var patPanel = tab1.add('panel', undefined, 'Preset Pattern');
        patPanel.orientation   = 'column';
        patPanel.alignChildren = 'fill';
        patPanel.margins       = [8, 12, 8, 6];

        var patRow = patPanel.add('group');
        patRow.add('statictext', undefined, 'Pattern:');
        var patDrop = patRow.add('dropdownlist', undefined, patternNames);
        patDrop.preferredSize.width = 260;
        var selIdx = 0;
        for (var pi = 0; pi < patternNames.length; pi++) {
            if (patternNames[pi] === settings.selectedPattern) { selIdx = pi; break; }
        }
        patDrop.selection = selIdx;

        var patRow2 = patPanel.add('group');
        var shuffleBtn = patRow2.add('button', undefined, '\u21BA  Random Pattern');
        shuffleBtn.preferredSize.width = 148;
        var shuffleStatus = patRow2.add('statictext', undefined, '');
        shuffleStatus.characters = 26;

        var _randomPattern = null;
        shuffleBtn.onClick = function() {
            _randomPattern = randomPattern();
            var rowCount = _randomPattern.length;
            var colCount = 0;
            for (var ri = 0; ri < _randomPattern.length; ri++) {
                var s = 0;
                for (var ci = 0; ci < _randomPattern[ri].length; ci++) s += (_randomPattern[ri][ci].cs||1);
                if (s > colCount) colCount = s;
            }
            shuffleStatus.text = 'Random ' + colCount + '\u00D7' + rowCount;
            updatePreview();
        };

        patDrop.onChange = function() { _randomPattern = null; shuffleStatus.text = ''; updatePreview(); };

        // Custom pattern
        var customPanel = tab1.add('panel', undefined, 'Custom Pattern');
        customPanel.orientation   = 'column';
        customPanel.alignChildren = 'fill';
        customPanel.margins       = [8, 12, 8, 6];

        var customRow = customPanel.add('group');
        var customBtn = customRow.add('button', undefined, 'Configure Custom\u2026');
        customBtn.preferredSize.width = 140;
        var customStatus = customRow.add('statictext', undefined, 'None active');
        customStatus.characters = 30;

        var _customPattern = null;
        customBtn.onClick = function() {
            var cd = new Window('dialog', 'Custom Pattern Input');
            cd.orientation = 'column'; cd.alignChildren = 'fill';
            cd.minimumSize.width = 460;

            var typeRow = cd.add('group');
            var simRad = typeRow.add('radiobutton', undefined, 'Simple  (2,1 | 1,1,1)');
            var jsonRad = typeRow.add('radiobutton', undefined, 'JSON (advanced)');
            simRad.value = true;

            var area = cd.add('panel', undefined, 'Simple Pattern');
            area.orientation = 'column'; area.alignChildren = 'fill'; area.margins = 10;
            area.preferredSize.height = 120;

            var simInput, jsonInput;

            var btnGrp = cd.add('group'); btnGrp.alignment = 'center';
            cd.add('button', undefined, 'Cancel').onClick = function() { cd.close(); };
            cd.add('button', undefined, 'Apply').onClick = function() {
                if (simRad.value) {
                    var pat = parseSimplePattern(simInput ? simInput.text : '');
                    if (pat && pat.length) {
                        _customPattern = pat;
                        _randomPattern = null; shuffleStatus.text = '';
                        customStatus.text = 'Simple: ' + (simInput.text || '');
                    }
                } else {
                    var pat2 = parseJSONPattern(jsonInput ? jsonInput.text : '');
                    if (pat2) {
                        _customPattern = pat2;
                        _randomPattern = null; shuffleStatus.text = '';
                        settings.customPatternJSON = jsonInput.text;
                        customStatus.text = 'JSON configured';
                    }
                }
                updatePreview();
                cd.close();
            };

            function rebuildArea() {
                while (area.children.length > 0) { try { area.remove(area.children[0]); } catch(_){break;} }
                if (simRad.value) {
                    area.text = 'Simple: rows by |, cols by ,';
                    simInput = area.add('edittext', undefined, settings.simplePattern||'');
                    simInput.characters = 42;
                    area.add('statictext', undefined, 'e.g.  2,1 | 1,1,1 | 3  generates 3 rows');
                } else {
                    area.text = 'JSON: Array of rows with {cs, rs} cells';
                    jsonInput = area.add('edittext', undefined, settings.customPatternJSON||'', {multiline:true});
                    jsonInput.preferredSize = [420, 80];
                    area.add('statictext', undefined, 'e.g. [[{"cs":2,"rs":2},{"cs":1,"rs":1}],[{"cs":1,"rs":1}]]');
                }
                cd.layout.layout(true);
            }
            simRad.onClick  = rebuildArea;
            jsonRad.onClick = rebuildArea;
            rebuildArea();
            cd.show();
        };

        // ════════════════════════════════════════════════════════════
        //  TAB 2 — LIVE PREVIEW
        // ════════════════════════════════════════════════════════════
        var tab2 = tabs.add('tab', undefined, ' \u25A3 Live Preview ');
        tab2.orientation   = 'column';
        tab2.alignChildren = 'fill';
        tab2.margins       = [8, 10, 8, 4];
        tab2.spacing       = 6;

        var prevPanel = tab2.add('panel', undefined, 'Grid Layout Preview');
        prevPanel.orientation   = 'column';
        prevPanel.alignChildren = 'fill';
        prevPanel.margins       = [8, 14, 8, 8];
        prevPanel.spacing       = 6;

        var previewText = prevPanel.add('edittext', undefined, '', {multiline: true, readonly: true});
        previewText.preferredSize = [490, 300];
        try { previewText.graphics.font = ScriptUI.newFont('Courier New', ScriptUI.FontStyle.REGULAR, 10); } catch(_) {}

        var previewInfo = prevPanel.add('statictext', undefined, '');
        try { previewInfo.graphics.font = ScriptUI.newFont('Arial', ScriptUI.FontStyle.ITALIC, 9); } catch(_) {}

        prevPanel.add('statictext', undefined, 'Switch to Grid & Pattern tab to change the layout.');

        // State vars
        var _previewPattern = bentoPatterns[patternNames[0]];
        var _previewFill    = [0.12, 0.12, 0.12];
        var _previewGutter  = 24;

        function buildPreviewText(pat) {
            if (!pat || !pat.length) return '(no pattern)';

            var ROWS = pat.length, COLS = 0;
            for (var ri = 0; ri < ROWS; ri++) {
                var s = 0;
                for (var ci = 0; ci < pat[ri].length; ci++) s += (pat[ri][ci].cs || 1);
                if (s > COLS) COLS = s;
            }
            if (COLS <= 0) return '(empty)';

            var CELL_W = Math.max(3, Math.floor(36 / COLS));
            var grid   = [];
            var occ    = {};
            var labels = [];
            var ri2, ci2, dr, dc;
            for (ri2 = 0; ri2 < ROWS; ri2++) {
                grid[ri2]   = [];
                labels[ri2] = [];
                for (ci2 = 0; ci2 < COLS; ci2++) {
                    grid[ri2][ci2]   = null;
                    labels[ri2][ci2] = null;
                }
            }

            function kp2(r,c){ return r+'_'+c; }
            var cellIdx = 0;
            for (ri2 = 0; ri2 < ROWS; ri2++) {
                var col2 = 0;
                for (var j2 = 0; j2 < pat[ri2].length; j2++) {
                    while (occ[kp2(ri2, col2)]) col2++;
                    var cell2 = pat[ri2][j2];
                    if (!cell2) { col2++; continue; }
                    var cs2 = cell2.cs || 1, rs2 = cell2.rs || 1;
                    var lbl = (cs2 > 1 || rs2 > 1) ? cs2+'x'+rs2 : '\u25A0';
                    labels[ri2][col2] = lbl;
                    for (dr = 0; dr < rs2; dr++) {
                        for (dc = 0; dc < cs2; dc++) {
                            occ[kp2(ri2+dr, col2+dc)] = cellIdx;
                            grid[ri2+dr][col2+dc] = cellIdx;
                        }
                    }
                    col2 += cs2;
                    cellIdx++;
                }
            }

            var hLine = '+';
            for (ci2 = 0; ci2 < COLS; ci2++) {
                for (var k=0;k<CELL_W;k++) hLine += '-';
                hLine += '+';
            }

            var lines = [hLine];
            for (ri2 = 0; ri2 < ROWS; ri2++) {
                var contentLine = '|';
                for (ci2 = 0; ci2 < COLS; ci2++) {
                    var lbl2 = labels[ri2][ci2] || '';
                    var pad = CELL_W - lbl2.length;
                    var lp  = Math.floor(pad/2), rp = pad - lp;
                    var cell3 = '';
                    for (var p=0;p<lp;p++) cell3 += (lbl2 ? ' ' : '\u00B7');
                    cell3 += lbl2;
                    for (var q=0;q<rp;q++) cell3 += (lbl2 ? ' ' : '\u00B7');
                    var rightBorder = '|';
                    if (ci2 < COLS-1 && grid[ri2][ci2] !== null && grid[ri2][ci2] === grid[ri2][ci2+1]) {
                        rightBorder = ' ';
                        cell3 = '';
                        for (var p2=0;p2<CELL_W;p2++) cell3 += ' ';
                    }
                    contentLine += cell3 + rightBorder;
                }
                lines.push(contentLine);

                var bLine = '+';
                for (ci2 = 0; ci2 < COLS; ci2++) {
                    var bottomChar = '-';
                    if (ri2 < ROWS-1 && grid[ri2][ci2] !== null && grid[ri2][ci2] === grid[ri2+1][ci2]) {
                        bottomChar = ' ';
                    }
                    for (var k2=0;k2<CELL_W;k2++) bLine += bottomChar;
                    var hSpan = (ci2 < COLS-1 && ri2 < ROWS-1 &&
                                 grid[ri2][ci2] !== null && grid[ri2][ci2] === grid[ri2][ci2+1] &&
                                 grid[ri2][ci2] === grid[ri2+1][ci2] && grid[ri2][ci2] === grid[ri2+1][ci2+1]);
                    bLine += hSpan ? ' ' : '+';
                }
                lines.push(bLine);
            }
            return lines.join('\n');
        }

        function updatePreview() {
            var pat = _customPattern || _randomPattern ||
                      (patDrop.selection ? bentoPatterns[patDrop.selection.text] : bentoPatterns[patternNames[0]]);
            _previewPattern = pat || bentoPatterns[patternNames[0]];
            _previewGutter  = parseFloat(gutterInput.text) || 24;
            _previewFill    = settings.fillColor || [0.12, 0.12, 0.12];

            previewText.text = buildPreviewText(_previewPattern);

            var ROWS3 = _previewPattern ? _previewPattern.length : 0;
            var COLS3 = 0;
            if (_previewPattern) {
                for (var ri3=0;ri3<ROWS3;ri3++) {
                    var s3=0;
                    for (var ci3=0;ci3<_previewPattern[ri3].length;ci3++) s3+=(_previewPattern[ri3][ci3].cs||1);
                    if(s3>COLS3) COLS3=s3;
                }
            }
            previewInfo.text = COLS3 + ' cols  \u00D7  ' + ROWS3 + ' rows   |   Gutter: ' + _previewGutter + 'px   |   ' +
                               Math.round(settings.gridWidth) + ' \u00D7 ' + Math.round(settings.gridHeight) + ' px';
        }

        // ════════════════════════════════════════════════════════════
        //  TAB 3 — COLOR & FILL
        // ════════════════════════════════════════════════════════════
        var tab3 = tabs.add('tab', undefined, ' \u2728 Color & Fill ');
        tab3.orientation   = 'column';
        tab3.alignChildren = 'fill';
        tab3.margins       = [8, 10, 8, 4];
        tab3.spacing       = 6;

        // ── Color Palette Presets ─────────────────────────────────────
        var palPanel = tab3.add('panel', undefined, 'Color Palette Presets');
        palPanel.orientation   = 'column';
        palPanel.alignChildren = 'fill';
        palPanel.margins       = 10;

        var palRow = palPanel.add('group');
        palRow.add('statictext', undefined, 'Palette:');
        var palDrop = palRow.add('dropdownlist', undefined, paletteNames);
        palDrop.preferredSize.width = 240;
        palDrop.selection = settings.paletteIndex || 0;

        var palHint = palPanel.add('statictext', undefined, 'Selects fill color from a curated palette set.');
        try { palHint.graphics.font = ScriptUI.newFont(palHint.graphics.font.name, ScriptUI.FontStyle.ITALIC, 10); } catch(_) {}

        // ── Fill Color ───────────────────────────────────────────────
        var fillPanel = tab3.add('panel', undefined, 'Fill Color');
        fillPanel.orientation   = 'column';
        fillPanel.alignChildren = 'fill';
        fillPanel.margins       = 10;

        var autoCheck = fillPanel.add('checkbox', undefined, 'Use Illustrator foreground fill  (auto-detect)');
        autoCheck.value = settings.useAutoFill;

        var fillRow = fillPanel.add('group');
        fillRow.add('statictext', undefined, 'Fill RGB (0\u20131):');
        var fR = fillRow.add('edittext', undefined, settings.fillColor[0].toFixed(2)); fR.characters = 5;
        var fG = fillRow.add('edittext', undefined, settings.fillColor[1].toFixed(2)); fG.characters = 5;
        var fB = fillRow.add('edittext', undefined, settings.fillColor[2].toFixed(2)); fB.characters = 5;
        var fillHexRow = fillPanel.add('group');
        fillHexRow.add('statictext', undefined, 'Hex:');
        var fillHexInput = fillHexRow.add('edittext', undefined, rgbToHex(settings.fillColor));
        fillHexInput.characters = 10;
        var fillHexBtn = fillHexRow.add('button', undefined, 'Apply Hex');
        fillHexBtn.preferredSize.width = 80;

        function setFillRGB(rgb) {
            fR.text = rgb[0].toFixed(2);
            fG.text = rgb[1].toFixed(2);
            fB.text = rgb[2].toFixed(2);
            fillHexInput.text = rgbToHex(rgb);
            settings.fillColor = rgb;
            updatePreview();
        }

        fillHexBtn.onClick = function() {
            var c = hexToRGB01(fillHexInput.text);
            if (c) setFillRGB(c);
            else alert('Invalid hex. Use format: #RRGGBB or RRGGBB');
        };

        fR.onChange = fG.onChange = fB.onChange = function() {
            var r = parseFloat(fR.text)||0, g = parseFloat(fG.text)||0, b = parseFloat(fB.text)||0;
            settings.fillColor = [clamp01(r), clamp01(g), clamp01(b)];
            fillHexInput.text  = rgbToHex(settings.fillColor);
            _previewFill = settings.fillColor;
            updatePreview();
        };

        palDrop.onChange = function() {
            var palName = paletteNames[palDrop.selection.index];
            var pal = colorPalettes[palName];
            if (pal && pal.length) { setFillRGB(pal[0]); }
            settings.paletteIndex = palDrop.selection.index;
        };

        function syncAutoFill() {
            var en = !autoCheck.value;
            fR.enabled = fG.enabled = fB.enabled = fillHexInput.enabled = fillHexBtn.enabled = en;
            palDrop.enabled = en;
            if (autoCheck.value) {
                var fg = getIllustratorForegroundRGB();
                if (fg) setFillRGB(fg);
            }
        }
        autoCheck.onClick = function() { settings.useAutoFill = autoCheck.value; syncAutoFill(); };
        syncAutoFill();

        // ════════════════════════════════════════════════════════════
        //  TAB 4 — STROKE & BACKGROUND
        // ════════════════════════════════════════════════════════════
        var tab4 = tabs.add('tab', undefined, ' \u25A1 Stroke & BG ');
        tab4.orientation   = 'column';
        tab4.alignChildren = 'fill';
        tab4.margins       = [8, 10, 8, 4];
        tab4.spacing       = 6;

        // ── Rounded Corners ──────────────────────────────────────────
        var cornPanel = tab4.add('panel', undefined, 'Rounded Corners');
        cornPanel.orientation   = 'column';
        cornPanel.alignChildren = 'fill';
        cornPanel.margins       = 10;

        var roundedCheck = cornPanel.add('checkbox', undefined, 'Enable Rounded Corners');
        roundedCheck.value = settings.useRoundedCorners;

        var cornRow = cornPanel.add('group');
        cornRow.add('statictext', undefined, 'Radius:');
        var cornInput = cornRow.add('edittext', undefined, String(settings.cornerRadius));
        cornInput.characters = 7;
        cornRow.add('statictext', undefined, 'px');
        cornInput.enabled = roundedCheck.value;
        roundedCheck.onClick = function() { cornInput.enabled = roundedCheck.value; };

        // ── Stroke ───────────────────────────────────────────────────
        var strokePanel = tab4.add('panel', undefined, 'Stroke');
        strokePanel.orientation   = 'column';
        strokePanel.alignChildren = 'fill';
        strokePanel.margins       = 10;

        var strokeCheck = strokePanel.add('checkbox', undefined, 'Add Stroke to cells');
        strokeCheck.value = settings.useStroke;

        var strokeRow = strokePanel.add('group');
        strokeRow.add('statictext', undefined, 'Width:');
        var strokeWInput = strokeRow.add('edittext', undefined, String(settings.strokeWidth));
        strokeWInput.characters = 4;
        strokeRow.add('statictext', undefined, 'px');

        var strokeColorRow = strokePanel.add('group');
        strokeColorRow.add('statictext', undefined, 'Stroke RGB (0\u20131):');
        var sR = strokeColorRow.add('edittext', undefined, settings.strokeColor[0].toFixed(2)); sR.characters = 4;
        var sG = strokeColorRow.add('edittext', undefined, settings.strokeColor[1].toFixed(2)); sG.characters = 4;
        var sB = strokeColorRow.add('edittext', undefined, settings.strokeColor[2].toFixed(2)); sB.characters = 4;

        var strokeHexRow = strokePanel.add('group');
        strokeHexRow.add('statictext', undefined, 'Stroke Hex:');
        var strokeHexInput = strokeHexRow.add('edittext', undefined, rgbToHex(settings.strokeColor));
        strokeHexInput.characters = 10;
        var strokeHexBtn = strokeHexRow.add('button', undefined, 'Apply');
        strokeHexBtn.preferredSize.width = 60;
        strokeHexBtn.onClick = function() {
            var c = hexToRGB01(strokeHexInput.text);
            if (c) { sR.text=c[0].toFixed(2); sG.text=c[1].toFixed(2); sB.text=c[2].toFixed(2); settings.strokeColor=c; }
        };

        function updateStrokeEnabled() {
            strokeWInput.enabled = sR.enabled = sG.enabled = sB.enabled = strokeHexInput.enabled = strokeHexBtn.enabled = strokeCheck.value;
        }
        strokeCheck.onClick = updateStrokeEnabled;
        updateStrokeEnabled();

        // ── Background Frame ─────────────────────────────────────────
        var bgPanel = tab4.add('panel', undefined, 'Background Frame');
        bgPanel.orientation   = 'column';
        bgPanel.alignChildren = 'fill';
        bgPanel.margins       = 10;

        var bgCheck = bgPanel.add('checkbox', undefined, 'Create Background Frame');
        bgCheck.value = settings.createBackground;

        var bgRow = bgPanel.add('group');
        bgRow.add('statictext', undefined, 'Padding:');
        var bgPadInput = bgRow.add('edittext', undefined, String(settings.backgroundPadding));
        bgPadInput.characters = 7;
        bgRow.add('statictext', undefined, 'px');

        var bgColorRow = bgPanel.add('group');
        bgColorRow.add('statictext', undefined, 'BG RGB (0\u20131):');
        var bgR = bgColorRow.add('edittext', undefined, settings.bgColor[0].toFixed(2)); bgR.characters = 4;
        var bgG = bgColorRow.add('edittext', undefined, settings.bgColor[1].toFixed(2)); bgG.characters = 4;
        var bgB = bgColorRow.add('edittext', undefined, settings.bgColor[2].toFixed(2)); bgB.characters = 4;

        function updateBGEnabled() {
            bgPadInput.enabled = bgR.enabled = bgG.enabled = bgB.enabled = bgCheck.value;
        }
        bgCheck.onClick = updateBGEnabled;
        updateBGEnabled();

        // ════════════════════════════════════════════════════════════
        //  TAB 5 — OPTIONS & SETTINGS
        // ════════════════════════════════════════════════════════════
        var tab5 = tabs.add('tab', undefined, ' \u2699 Options ');
        tab5.orientation   = 'column';
        tab5.alignChildren = 'fill';
        tab5.margins       = [8, 10, 8, 4];
        tab5.spacing       = 6;

        // Naming options
        var namePanel = tab5.add('panel', undefined, 'Naming');
        namePanel.orientation   = 'column';
        namePanel.alignChildren = 'fill';
        namePanel.margins       = 10;

        var nameRow = namePanel.add('group');
        nameRow.add('statictext', undefined, 'Group name:');
        var groupNameInput = nameRow.add('edittext', undefined, settings.groupName);
        groupNameInput.characters = 24;

        var prefRow = namePanel.add('group');
        prefRow.add('statictext', undefined, 'Cell prefix:');
        var cellPrefixInput = prefRow.add('edittext', undefined, settings.namingPrefix);
        cellPrefixInput.characters = 16;
        namePanel.add('statictext', undefined, 'Cells are named Prefix_Row_Col  (e.g. Cell_0_0).');

        // Multi-color
        var multiPanel = tab5.add('panel', undefined, 'Multi-Color Cells  (from Palette)');
        multiPanel.orientation   = 'column';
        multiPanel.alignChildren = 'fill';
        multiPanel.margins       = 10;

        var multiCheck = multiPanel.add('checkbox', undefined, 'Cycle through palette colors across cells');
        multiCheck.value = false;
        var multiNote = multiPanel.add('statictext', undefined, 'When enabled, each cell steps through the 3 palette shades in order.');
        multiNote.properties = {multiline: true};
        try { multiNote.graphics.font = ScriptUI.newFont(multiNote.graphics.font.name, ScriptUI.FontStyle.ITALIC, 10); } catch(_) {}

        // Save / Load
        var savePanel = tab5.add('panel', undefined, 'Save \u2215 Load Settings');
        savePanel.orientation   = 'column';
        savePanel.alignChildren = 'fill';
        savePanel.margins       = 10;

        var saveBtnRow = savePanel.add('group');
        var saveBtn = saveBtnRow.add('button', undefined, '\u2193  Save Settings');
        saveBtn.preferredSize.width = 130;
        var loadBtn = saveBtnRow.add('button', undefined, '\u2191  Load Settings');
        loadBtn.preferredSize.width = 130;

        var saveHint = savePanel.add('statictext', undefined, 'Settings saved to: ' + SETTINGS_FILE_PATH);
        saveHint.properties = {multiline: true};
        try { saveHint.graphics.font = ScriptUI.newFont(saveHint.graphics.font.name, ScriptUI.FontStyle.ITALIC, 9); } catch(_) {}

        saveBtn.onClick = function() {
            collectSettings();
            saveSettings();
            alert('Settings saved successfully!\n' + SETTINGS_FILE_PATH);
        };
        loadBtn.onClick = function() {
            loadSettings();
            alert('Settings loaded. Please re-open the dialog to see them applied.');
        };

        // ════════════════════════════════════════════════════════════
        //  TAB 6 — ABOUT & LINKS
        // ════════════════════════════════════════════════════════════
        var tab6 = tabs.add('tab', undefined, ' \u2665 About ');
        tab6.orientation   = 'column';
        tab6.alignChildren = 'fill';
        tab6.margins       = [12, 12, 12, 8];
        tab6.spacing       = 8;

        // Bio
        var bioPanel = tab6.add('panel', undefined, 'Shirish Salve');
        bioPanel.orientation   = 'column';
        bioPanel.alignChildren = 'fill';
        bioPanel.margins       = [12, 14, 12, 10];
        bioPanel.spacing       = 6;

        var bioText = bioPanel.add('statictext', undefined,
            'Pune-based Graphic Designer  \u2022  12+ Years Experience\n' +
            'Logo Design  \u2022  UI/UX  \u2022  Illustration  \u2022  Art',
            {multiline: true});
        try { bioText.graphics.font = ScriptUI.newFont('Arial', ScriptUI.FontStyle.REGULAR, 11); } catch(_) {}

        var tagLine = bioPanel.add('statictext', undefined, 'Proficient in nearly every design tool. Let\'s create something amazing!');
        try { tagLine.graphics.font = ScriptUI.newFont('Arial', ScriptUI.FontStyle.ITALIC, 10); } catch(_) {}

        // Links
        var linksPanel = tab6.add('panel', undefined, 'Connect');
        linksPanel.orientation   = 'column';
        linksPanel.alignChildren = 'fill';
        linksPanel.margins       = [12, 14, 12, 10];
        linksPanel.spacing       = 8;

        // ── Behance ──────────────────────────────────────────────────
        var behanceRow = linksPanel.add('group');
        behanceRow.add('statictext', undefined, '\u25B6  Behance:');
        var behanceField = behanceRow.add('edittext', undefined, 'https://www.behance.net/myselfshirish');
        behanceField.preferredSize.width = 260;
        var behanceBtn = behanceRow.add('button', undefined, 'Open');
        behanceBtn.preferredSize.width = 54;
        behanceBtn.onClick = function() {
            try {
                app.openURL('https://www.behance.net/myselfshirish');
            } catch(e) {
                behanceField.active = true;
                alert('Could not open automatically.\nURL is selected in the field above — copy it and paste in your browser.');
            }
        };

        // Donation
        var donatePanel = tab6.add('panel', undefined, 'Support This Script  \u2615');
        donatePanel.orientation   = 'column';
        donatePanel.alignChildren = 'fill';
        donatePanel.margins       = [12, 14, 12, 10];
        donatePanel.spacing       = 6;

        donatePanel.add('statictext', undefined, 'If this script saved you time, a coffee would mean the world! \u2665');

        var donateRow = donatePanel.add('group');
        donateRow.alignment = 'fill';
        donateRow.alignChildren = 'fill';
        var donateField = donateRow.add('edittext', undefined, 'https://buymeacoffee.com/shirish');
        donateField.preferredSize.width = 280;
        var donateBtn = donateRow.add('button', undefined, '\u2615 Open');
        donateBtn.preferredSize.width = 70;
        donateBtn.onClick = function() {
            try {
                app.openURL('https://buymeacoffee.com/shirish');
            } catch(e) {
                donateField.active = true;
                alert('Could not open automatically.\nURL is selected in the field above — copy it and paste in your browser.');
            }
        };

        var donateTip = donatePanel.add('statictext', undefined, 'Click the field, Ctrl+A to select all, then Ctrl+C to copy.');
        try { donateTip.graphics.font = ScriptUI.newFont('Arial', ScriptUI.FontStyle.ITALIC, 9); } catch(_) {}

        // Script info
        var scriptInfoPanel = tab6.add('panel', undefined, 'Script Info');
        scriptInfoPanel.orientation   = 'column';
        scriptInfoPanel.alignChildren = 'fill';
        scriptInfoPanel.margins       = [10, 12, 10, 8];
        scriptInfoPanel.spacing       = 4;

        scriptInfoPanel.add('statictext', undefined, 'BentoGrid v3.0  \u2014  Adobe Illustrator ExtendScript');
        scriptInfoPanel.add('statictext', undefined, '50+ patterns  \u2022  20 palette presets  \u2022  Live preview');
        scriptInfoPanel.add('statictext', undefined, '\u00A9 2026 Shirish Salve  \u2014  All rights reserved.');

        // ── ACTION BUTTONS ────────────────────────────────────────────
        var actPanel = dlg.add('group');
        actPanel.alignment   = 'fill';
        actPanel.orientation = 'row';
        actPanel.alignChildren = 'center';
        actPanel.margins     = [10, 6, 10, 6];
        actPanel.spacing     = 10;

        var cancelBtn = actPanel.add('button', undefined, 'Cancel');
        cancelBtn.preferredSize.width = 90;
        cancelBtn.onClick = function() { dlg.close(); };

        var generateBtn = actPanel.add('button', undefined, '  Generate Bento Grid  \u25BA');
        generateBtn.preferredSize.width = 200;

        // ── COLLECT SETTINGS ─────────────────────────────────────────
        function collectSettings() {
            settings.gridWidth          = parseInt(widthInput.text,  10) || 1200;
            settings.gridHeight         = parseInt(heightInput.text, 10) || 600;
            settings.gutter             = parseFloat(gutterInput.text)   || 0;
            settings.useRoundedCorners  = roundedCheck.value;
            settings.cornerRadius       = parseFloat(cornInput.text)     || 16;
            settings.useStroke          = strokeCheck.value;
            settings.strokeWidth        = parseFloat(strokeWInput.text)  || 1;
            settings.strokeColor        = [parseFloat(sR.text)||1, parseFloat(sG.text)||1, parseFloat(sB.text)||1];
            settings.createBackground   = bgCheck.value;
            settings.backgroundPadding  = parseFloat(bgPadInput.text)   || 24;
            settings.bgColor            = [parseFloat(bgR.text)||0.95, parseFloat(bgG.text)||0.95, parseFloat(bgB.text)||0.95];
            settings.groupName          = groupNameInput.text  || 'Bento Grid';
            settings.namingPrefix       = cellPrefixInput.text || 'Cell';
            settings.paletteIndex       = palDrop.selection ? palDrop.selection.index : 0;
            settings.useMultiColor      = multiCheck.value;

            // Fill color
            if (settings.useAutoFill) {
                var fg = getIllustratorForegroundRGB();
                settings.fillColor = fg || [0.12, 0.12, 0.12];
            } else {
                settings.fillColor = [clamp01(parseFloat(fR.text)||0.12), clamp01(parseFloat(fG.text)||0.12), clamp01(parseFloat(fB.text)||0.12)];
            }

            // Active palette for multi-color
            var palName = paletteNames[settings.paletteIndex];
            settings.activePalette = colorPalettes[palName];

            // Pattern selection
            if (_customPattern) {
                settings.resolvedPattern = _customPattern;
                settings.selectedPattern = 'Custom';
            } else if (_randomPattern) {
                settings.resolvedPattern = _randomPattern;
                settings.selectedPattern = 'Random';
            } else {
                var patName = patDrop.selection ? patDrop.selection.text : patternNames[0];
                settings.resolvedPattern = bentoPatterns[patName];
                settings.selectedPattern = patName;
            }
        }

        generateBtn.onClick = function() {
            collectSettings();

            if (!(settings.gridWidth > 0 && settings.gridHeight > 0)) { alert('Grid size must be positive.'); return; }
            if (!(settings.gutter >= 0))                               { alert('Gutter must be 0 or positive.'); return; }
            if (!settings.resolvedPattern || !settings.resolvedPattern.length) { alert('No valid pattern found.'); return; }

            dlg.close();
            generateBentoGrid();
        };

        // Trigger initial preview render
        updatePreview();

        // Tight dialog — logo + tabs + buttons always visible
        dlg.preferredSize  = [560, 590];
        dlg.maximumSize    = [560, 590];
        dlg.layout.layout(true);
        dlg.layout.resize();
        return dlg;
    }

    // ================================================================
    //  GRID GENERATION
    // ================================================================
    function cellKey(r, c) { return r + '_' + c; }

    function getMultiColor(idx) {
        if (!settings.useMultiColor || !settings.activePalette || !settings.activePalette.length) {
            return settings.fillColor;
        }
        return settings.activePalette[idx % settings.activePalette.length];
    }

    function createGridCell(group, x, y, w, h, row, col, cellIndex) {
        var it;
        var cr = (settings.useRoundedCorners && settings.cornerRadius > 0) ? settings.cornerRadius : 0;
        if (cr > 0) {
            it = group.pathItems.roundedRectangle(y, x, w, h, cr, cr);
        } else {
            it = group.pathItems.rectangle(y, x, w, h);
        }
        it.filled    = true;
        it.fillColor = createColor(getMultiColor(cellIndex));
        it.stroked   = settings.useStroke;
        if (settings.useStroke) {
            it.strokeColor = createColor(settings.strokeColor);
            it.strokeWidth = settings.strokeWidth;
        }
        it.name = settings.namingPrefix + '_' + row + '_' + col;
        return it;
    }

    function generateSpanningGrid(group, startX, startY, pattern) {
        var rows = pattern.length, cols = 0;
        for (var r = 0; r < rows; r++) {
            var sum = 0;
            for (var j = 0; j < pattern[r].length; j++) sum += (pattern[r][j].cs || 1);
            if (sum > cols) cols = sum;
        }
        if (cols === 0) return;

        var cellW = (settings.gridWidth  - settings.gutter * (cols - 1)) / cols;
        var cellH = (settings.gridHeight - settings.gutter * (rows - 1)) / rows;
        var occ   = {};
        var cellIdx = 0;

        for (var r2 = 0; r2 < rows; r2++) {
            var c = 0;
            for (var k = 0; k < pattern[r2].length; k++) {
                while (occ[cellKey(r2, c)]) c++;
                var cell = pattern[r2][k];
                if (!cell) { c++; continue; }
                var cs = cell.cs || 1, rs = cell.rs || 1;
                var x  = startX + c * (cellW + settings.gutter);
                var y  = startY - r2 * (cellH + settings.gutter);
                var w  = cellW * cs + settings.gutter * (cs - 1);
                var h  = cellH * rs + settings.gutter * (rs - 1);
                createGridCell(group, x, y, w, h, r2, c, cellIdx++);
                for (var dr = 0; dr < rs; dr++)
                    for (var dc = 0; dc < cs; dc++)
                        occ[cellKey(r2 + dr, c + dc)] = true;
                c += cs;
            }
        }
    }

    function createBackgroundFrame(group, startX, startY) {
        var pad = settings.backgroundPadding;
        var x   = startX - pad,  y  = startY + pad;
        var w   = settings.gridWidth  + pad * 2;
        var h   = settings.gridHeight + pad * 2;
        var cr  = (settings.useRoundedCorners && settings.cornerRadius > 0) ? settings.cornerRadius * 1.5 : 0;
        var bg  = (cr > 0)
            ? group.pathItems.roundedRectangle(y, x, w, h, cr, cr)
            : group.pathItems.rectangle(y, x, w, h);
        bg.filled    = true;
        bg.fillColor = createColor(settings.bgColor);
        bg.stroked   = settings.useStroke;
        if (settings.useStroke) { bg.strokeColor = createColor(settings.strokeColor); bg.strokeWidth = settings.strokeWidth; }
        bg.name = 'Background';
        try { bg.zOrder(ZOrderMethod.SENDTOBACK); } catch(e) {}
    }

    function generateBentoGrid() {
        var ab     = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var b      = ab.artboardRect;
        var startX = b[0], startY = b[1];

        var group  = doc.groupItems.add();
        group.name = settings.groupName;

        if (settings.createBackground) createBackgroundFrame(group, startX, startY);

        generateSpanningGrid(group, startX, startY, settings.resolvedPattern);

        // Centre group on artboard
        var abCX = b[0] + (b[2] - b[0]) / 2;
        var abCY = b[1] + (b[3] - b[1]) / 2;
        var gb   = group.visibleBounds;
        var gCX  = (gb[0] + gb[2]) / 2;
        var gCY  = (gb[1] + gb[3]) / 2;
        group.left += (abCX - gCX);
        group.top  += (abCY - gCY);

        doc.selection = null;
        group.selected = true;

        // Auto-save settings after each successful generation
        saveSettings();
    }

    // ── Run ──────────────────────────────────────────────────────────
    var dlg = createDialog();
    dlg.show();

})();
