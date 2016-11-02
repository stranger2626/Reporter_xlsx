var fs = require('fs');
var path = require('canonical-path');
var _ = require('lodash');
var xl = require('excel4node');

// Custom reporter
var reporter_xlsx = function(options) {

    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1');

    var _defaultOutputFile = path.resolve(process.cwd(), './_test-output', 'excel.xlsx');
    options.outputFile = options.outputFile || _defaultOutputFile;

    initOutputFile(options.outputFile);
    options.appDir = options.appDir ||  './';
    var _root = { appDir: options.appDir, suites: [] };
    log('AppDir: ' + options.appDir, +1);
    var _currentSuite;

    this.suiteStarted = function(suite) {
        _currentSuite = { description: suite.description, status: null, specs: [] };
        _root.suites.push(_currentSuite);
        log('Suite: ' + suite.description, +1);
    };

    this.suiteDone = function(suite) {
        var statuses = _currentSuite.specs.map(function(spec) {
            return spec.status;
        });
        statuses = _.uniq(statuses);
        var status = statuses.indexOf('Failed') >= 0 ? 'Failed' : statuses.join(', ');
        _currentSuite.status = status;
        log('Suite ' + _currentSuite.status + ': ' + suite.description, -1);
    };

    this.specStarted = function(spec) {

    };

    this.specDone = function(spec) {
        var currentSpec = {
            description: spec.description,
            status: spec.statuses
        };
        if (spec.failedExpectations.length > 0) {
            currentSpec.failedExpectations = spec.failedExpectations;
        }

        _currentSuite.specs.push(currentSpec);
        log(spec.status + ' - ' + spec.description);
    };

    this.jasmineDone = function() {
        outputFile = options.outputFile;
        var output = formatOutput(_root);
        wb.write(options.outputFile);
    };

    function formatOutput(output) {
        var indent = '  ';
        var pad = '  ';
        var results = [];
        ws.cell(2,1).string('AppDir:' + output.appDir);
        var i = 3;
        output.suites.forEach(function(suite) {
            ws.cell(2,1).string(pad + 'Suite: ' + suite.description + ' -- ' + suite.status);
            i++;

        });
        return wb.write(options.outputFile);
    }
    function ensureDirectoryExistence(filePath) {
        var dirname = path.dirname(filePath);
        if (directoryExists(dirname)) {
            return true;
        }
        ensureDirectoryExistence(dirname);
        fs.mkdirSync(dirname);
    }

    function directoryExists(path) {
        try {
            return fs.statSync(path).isDirectory();
        }
        catch (err) {
            return false;
        }
    }

    function initOutputFile(outputFile) {
        ensureDirectoryExistence(outputFile);
        var header = "Protractor results for: " + (new Date()).toLocaleString() + "\n\n";
        ws.cell(1,1).string(header);
        wb.write(options.outputFile);
    }
    // for console output
    var _pad;
    function log(str, indent) {
        _pad = _pad || '';
        if (indent == -1) {
            _pad = _pad.substr(2);
        }
        console.log(_pad + str);
        if (indent == 1) {
            _pad = _pad + '  ';
        }
    }
};

module.exports = reporter_xlsx;

