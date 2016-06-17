/**
 * Created by fantastik on 16/06/2016.
 */

function addEventHandler(obj, evt, handler) {
    if(obj.addEventListener) {
        // W3C method
        obj.addEventListener(evt, handler, false);
    } else if(obj.attachEvent) {
        // IE method.
        obj.attachEvent('on'+evt, handler);
    } else {
        // Old school method.
        obj['on'+evt] = handler;
    }
}

Function.prototype.bindToEventHandler = function bindToEventHandler() {
    var handler = this;
    var boundParameters = Array.prototype.slice.call(arguments);
    //create closure
    return function(e) {
        e = e || window.event; // get window.event if e argument missing (in IE)
        boundParameters.unshift(e);
        handler.apply(this, boundParameters);
    }
};

function parseXlsxTreeStructure(book) {
    var sheet = book.Sheets[book.SheetNames[1]];
    var rawList = parseInitialXlsxList(sheet);
    var tree = createTreeStruct(rawList);
    saveWorkbook(tree);
}

function parseInitialXlsxList(sheet) {
    var dict = {};
    var rowNum = 2;
    while (sheet['A' + rowNum]) {
        var a = sheet['A' + rowNum];
        var b = sheet['B' + rowNum];
        var d = sheet['D' + rowNum];

        var id = a.w;
        var name = b.w
        var parent = d.w;

        dict[id] = {id:id, name:name, parent:parent, children:{}}

        rowNum++;
    }

    return dict;
}

function createTreeStruct(list) {
    var tree = {};
    for (l in list) {
        addNodeToTree(tree, list[l], list);
    }
    return tree;
}

function addNodeToTree(tree, node, list) {
    if (node.parent == "") {
        tree[node.id] = node;
        return;
    }

    // search root category
    var path = [];
    var n = node;
    do {
        path.splice(0, 0, n);
        n = list[n.parent];
    } while (n.parent != "");
    path.splice(0, 0, n);

    // for safety try to insert root
    addNodeToTree(tree, n, list);

    for (var i = 0, end = path.length - 1; i < end; i++) {
        var parent = path[i];
        var child = path[i + 1];

        if (parent.children[child.id] === undefined) {
            parent.children[child.id] = child;
        }
    }
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function sheetFromTree(tree) {
    var ws = {};
    var range = {s: {c:0, r:0}, e: {c:20, r:0 }};

    ws[XLSX.utils.encode_cell({c:0, r:0})] = makeCell("Category");

    var row = 1;
    for(var R in tree) {
        // Category
        var n = tree[R];

        ws[XLSX.utils.encode_cell({c:0, r:row})] = makeCell(n.name);

        for (var sub in n.children) {
            row += extractSubcategory(ws, 1, row, n.children[sub]);
        }
        row++;
    }
    range.e.r = row + 1;

    ws['!ref'] = XLSX.utils.encode_range(range);

    return ws;
}

function makeCell(val) {
    var cell = {v: val };

    if(typeof cell.v === 'number') cell.t = 'n';
    else if(typeof cell.v === 'boolean') cell.t = 'b';
    else if(cell.v instanceof Date) {
        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
    }
    else cell.t = 's';

    return cell;
}

function extractSubcategory(ws, col, row, node) {
    ws[XLSX.utils.encode_cell({c:col, r:0})] = makeCell("Sub-category");
    var r = row
    ws[XLSX.utils.encode_cell({c:col, r:row})] = makeCell(node.name);

    for (var sub in node.children) {
        row += extractSubcategory(ws, col + 1, row, node.children[sub]);
        row++;
    }

    return row - r;
}

function saveWorkbook(tree) {
    var wb = new Workbook();
    var ws = sheetFromTree(tree);
    /* add worksheet to workbook */
    wb.SheetNames.push("SEO Site structure");
    wb.Sheets["SEO Site structure"] = ws;
    var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});

    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "output.xlsx");
}

(function (){
    if(window.FileReader) {
        addEventHandler(window, 'load', function() {
            var status = document.getElementById('status');
            var drop   = document.getElementById('drop');
            var list   = document.getElementById('list');

            function cancel(e) {
                if (e.preventDefault) { e.preventDefault(); }
                return false;
            }

            // Tells the browser that we *can* drop on this target
            addEventHandler(drop, 'dragover', cancel);
            addEventHandler(drop, 'dragenter', cancel);
            addEventHandler(drop, 'drop', function (e) {
                e = e || window.event; // get window.event if e argument missing (in IE)
                if (e.preventDefault) { e.preventDefault(); } // stops the browser from redirecting off to the image.

                var dt    = e.dataTransfer;
                var files = dt.files;
                for (var i=0; i<files.length; i++) {
                    var file = files[i];
                    var reader = new FileReader();

                    addEventHandler(reader, 'loadend', function(e, file) {
                        var newFile       = document.createElement('div');
                        newFile.innerHTML = 'Loaded : '+file.name+' size '+file.size+' B';
                        list.appendChild(newFile);
                        var fileNumber = list.getElementsByTagName('div').length;
                        status.innerHTML = fileNumber < files.length
                            ? 'Loaded 100% of file '+fileNumber+' of '+files.length+'...'
                            : 'Done loading. processed '+fileNumber+' files.';


                        var workbook = XLSX.read(e.target.result, {type: 'binary'});
                        parseXlsxTreeStructure(workbook);
                    }.bindToEventHandler(file));

                    reader.readAsBinaryString(file);
                }
                return false;
            });
        });
    } else {
        document.getElementById('status').innerHTML = 'Your browser does not support the HTML5 FileReader.';
    }
})();