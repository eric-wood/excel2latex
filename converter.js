!(function($) {
  "use strict";

  // Global variable with information on current workbooks. Don't judge me.
  var workbooks = {};

  // Another global to store LaTeX output. Don't judge me...again...
  var latexOutput = {};

  var excelParser = {
    //latexEnvironment: 'tabular',
    latexEnvironment: 'array',

    latexEscape: function(text) {
      if(!$('#escape').is(':checked')) return text;

      var escapeRegExpr = function(str) {
        return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
      };

      var specials = ['\\', '&', '%', '$', '#', '_', '{', '}', '~', '^'];
      $.each(specials, function() {
        var regexp = new RegExp(escapeRegExpr(this), 'g');
        text = text.replace(regexp, '\\' + this, 'g');
      });

      return text;
    },

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
      for(i=0; i < numCols; i++) {
        args[i] = 'l';
      }
      args = ' | ' + args.join(' | ') + ' | ';
      var latex = "\\begin{" + excelParser.latexEnvironment + "}{" + args + "}\n\\hline\n";
      for(i=0; i < table.length; i++) {
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

      latex += "\\end{" + excelParser.latexEnvironment + "}\n";

      return latex;
    },

    processSheet: function(data, stringTable) {
      // get jQuery object out of the data for accessing stuffs
      var doc = $(data);

      var table = [];

      var rows = doc.find('sheetdata row');
      $.each(rows, function(i,row) {
        var rowNum = parseInt($(row).attr('r'), 10);

        // get columns
        var cols = $(row).find('c');
        var colVals = $.map(cols, function(col,j) {
          col = $(col);
          var val = excelParser.latexEscape(col.find('v').text());
          if(col.attr('t') == 's') {
            return excelParser.latexEscape(stringTable[parseInt(val, 10)]);
          } else {
            return val;
          }
        });
        table[rowNum-1] = colVals;
      });

      return table;
    },

    handleSheets: function(entries, stringTable) {
      // get the workbook.xml file, which contains the names of all workbooks
      var workbookMeta = entries.filter(function(entry) {
        return entry.filename === 'xl/workbook.xml';
      })[0];

      // filter out all files that aren't worksheets
      var sheets = $.grep(entries, function(n,i) {
        var filter1 = /^xl\/worksheets\/.*\.xml$/;
        var filter2 = /^xl\/worksheets\/_rels/;
        return (filter1.test(n.filename)) && (!filter2.test(n.filename));
      });

      // read the workbook meta data to get the names and crap
      workbookMeta.getData(new zip.TextWriter(), function(text) {
        var doc = $(text);

        // extract the names of the workbooks and their IDs for use later on...
        $.each(doc.find('sheets sheet'), function(i, tag) {
          tag = $(tag);
          var id = tag.attr('sheetId');
          var name = tag.attr('name');
          workbooks[id] = name;
        });

        excelParser.updateSelect();

        // iterate over all sheets and convert them to LaTeX
        $.each(sheets, function(_, sheet) {
          // the ID of the spreadsheet can only be found in the filename apparently :P
          var id = sheet.filename.match(/(\d)\.xml/)[1];

          sheet.getData(new zip.TextWriter(), function(text) {
            var table = excelParser.processSheet(text, stringTable);
            var latex = excelParser.toLatex(table);
            latexOutput[id] = latex;

            // I apologize for the hack :(
            if(id === '1') {
              excelParser.showOutput(1);
            }
          });
        });
      });
    },

    showOutput: function(id) {
      var latex = latexOutput[id];
      $('#latex-output').val(latex);
      $('#preview').html('$$\n' + latex + '\n$$');
      MathJax.Hub.Typeset("preview");
    },

    handleFiles: function(event) {
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
            var regexp = /^xl\/sharedStrings\.xml$/;
            return regexp.test(n.filename);
          })[0];

          stFile.getData(new zip.TextWriter(), function(text) {
            var stringTable = excelParser.parseStringTable(text);
            excelParser.handleSheets(entries, stringTable);
          });

          return undefined;
        });
      });

      return undefined;
    },

    updateSelect: function() {
      var select = $('#workbook');

      // clear existing option tags
      select.empty();

      for(var id in workbooks) {
        var tag = $('<option value="' + id + '">' + workbooks[id] + '</option>');
        tag.appendTo(select);
      }
    },

    init: function() {
      // hack because of jQuery shenanigans
      jQuery.event.props.push('dataTransfer');

      $('body').bind('dragover', function(event) {
        event.stopPropagation();
        event.preventDefault();
        event.dataTransfer.dropEffect = 'copy';
      });

      $('body').bind('drop', excelParser.handleFiles);

      // when a new workbook is selected, do stuff!
      $('#workbook').change(function(event) {
        var select = $(event.target);
        excelParser.showOutput(select.val());
      });
    }
  };

  $(function() {
    zip.workerScriptsPath = 'zip/';
    excelParser.init();
  });
})(window.jQuery);
