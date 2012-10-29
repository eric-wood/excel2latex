excelParser = {
  parseStringTable: function(data) {
    var doc = $(data);
    var stringTags = doc.find('si');
    var strings = $.map(stringTags, function(s,i) {
      return $(s).find('t').text();
    });

    return strings;
  },

  toLatex: function(table) {
    var max = 0;
    for(var i=0; i < table.length; i++) {
      if(table[i] && table[i].length > max) { max = table[i].length; }
    }

    var numCols = max;
    var args = [];
    for(var i=0; i < numCols; i++) {
      args[i] = 'l';
    }
    args = ' | ' + args.join(' | ') + ' | ';
    var latex = "\\begin{tabular}{" + args + "}\n\\hline\n";
    for(var i=0; i < table.length; i++) {
      var cols = table[i];
      // TODO: replace "&" with "\&"
      if(cols === undefined) { cols = []; }
      if(cols.length < numCols) {
        for(var x=cols.length; x < max; x++) {
          cols[x] = '\\ ';
        }
      }

      latex += "\t" + cols.join(' & ');
      latex += " \\\\ \\hline\n";
    }

    latex += "\\end{tabular}\n";
    
    $('#latex-output').val(latex);
  },

  processSheet: function(data, stringTable) {
    // get jQuery object out of the data for accessing stuffs
    var doc = $(data);

    var table = [];

    var rows = doc.find('sheetdata row');
    $.each(rows, function(i,row) {
      var rowNum = parseInt($(row).attr('r'));

      // get columns
      var cols = $(row).find('c');
      var colVals = $.map(cols, function(col,j) {
        var col = $(col);
        var val = col.find('v').text();
        if(col.attr('t') == 's') {
          return stringTable[parseInt(val)];
        } else {
          return val;
        }
      });
      table[rowNum-1] = colVals;
    });

    return table;
  },

  handleSheet: function(entries, stringTable) {
    // filter out all files that aren't worksheets
    var sheets = $.grep(entries, function(n,i) {
      return /^xl\/worksheets\//.test(n.filename);
    });

    // for now, only process the first sheet
    // TODO: give the user a choice of which sheet to process
    var sheet = sheets[0];
    sheet.getData(new zip.TextWriter(), function(text) {
      var table = excelParser.processSheet(text, stringTable);
      excelParser.toLatex(table);
    });
  },

  handleFile: function(event) {
    // prevent default browser behavior
    event.stopPropagation();
    event.preventDefault();

    // get file information
    var files = event.dataTransfer.files;
    if(!files.length) {
      // no files were given somehow...
      $('#latex-output').val('No files were given...try again?');
      return false;
    }

    var blob = files[0];

    // unzip and process files
    zip.createReader(new zip.BlobReader(blob), function(reader) {
      reader.getEntries(function(entries) {
        if(!entries.length) { return false; }

        // get the string table
        var stFile = $.grep(entries, function(n,i) {
          return /^xl\/sharedStrings\.xml$/.test(n.filename);
        })[0];

        stFile.getData(new zip.TextWriter(), function(text) {
          var stringTable = excelParser.parseStringTable(text);
          excelParser.handleSheet(entries, stringTable);
        });

      });
    });
  },

  init: function() {
    // hack because of jQuery shenanigans
    jQuery.event.props.push('dataTransfer');

    $('body').bind('dragover', function(event) {
      event.stopPropagation();
      event.preventDefault();
      event.dataTransfer.dropEffect = 'copy';
    });

    $('body').bind('drop', excelParser.handleFile);
  }
};

$(function() {
  zip.workerScriptsPath = 'zip/';
  excelParser.init();
});
