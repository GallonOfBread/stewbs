// Uno deck color schemes, shared by the Spicy Uno pages on stewbs.com.
//
// Load this with a plain <script> in <head>. It sets window.UNO_DECKS.
//
// Each scheme has:
//   a–d  the four suit colors, as [name, CSS color]
//   h    the suit used as the page highlight (rulebook)
//   f    flame colors for the logo, as suit slots (rulebook)
//   t    logo text colors, as suit slots (rulebook)
window.UNO_DECKS = {
    "70s": {
        a: ["Banana", "#D9B348"],
        b: ["Berry",  "#57989D"],
        c: ["Orange", "#C4602C"],
        d: ["Plum",   "#A62947"],
        h: "a",
        f: ["b", "d", "c", "a"],
        t: ["c", "b", "a", "d"]
    },
    "80s": {
        a: ["Blue",   "#0e9ad0"],
        b: ["Green",  "#5a9b64"],
        c: ["Orange", "#d4652c"],
        d: ["Purple", "#965490"],
        h: "a",
        f: ["d", "c", "b", "a"],
        t: ["d", "c", "a", "b"]
    },
    "90s": {
        a: ["Gold",   "#FEB81B"],
        b: ["Pink",   "#DA80B5"],
        c: ["Purple", "#5863B1"],
        d: ["Teal",   "#50C1C4"],
        h: "a",
        f: ["d", "c", "b", "a"],
        t: ["d", "c", "a", "b"]
    },
    "00s": {
        a: ["Blue",   "#5b9bc6"],
        b: ["Green",  "#497b5a"],
        c: ["Red",    "#c62d28"],
        d: ["Yellow", "#e4e339"],
        h: "a",
        f: ["a", "b", "d", "c"],
        t: ["a", "b", "d", "c"]
    },
    "tmnt": {
        a: ["Leo",    "#0269C6"],
        b: ["Donnie", "#7F3D97"],
        c: ["Raph",   "#E10E17"],
        d: ["Mikey",  "#F39904"],
        h: "a",
        f: ["a", "b", "c", "d"],
        t: ["d", "c", "a", "b"]
    },
    "default": {
        a: ["Blue",   "Blue"  ],
        b: ["Green",  "Green" ],
        c: ["Red",    "Red"   ],
        d: ["Yellow", "Yellow"],
        h: "c",
        f: ["a", "b", "d", "c"],
        t: ["a", "b", "d", "c"]
    }
};
