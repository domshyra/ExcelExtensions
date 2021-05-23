import React from 'react';
import PropTypes from 'prop-types';

const ParseExceptions = (props) => {
    let index = 1;
    const items = props.json.map((exception) => <ParseException exception={exception} key={index++} />);
    return (
        <div>
            <ul className="list-group">{items}</ul>
        </div>
    );
};
ParseExceptions.propTypes = {
    json: PropTypes.array,
};
export default ParseExceptions;

class Exception {
    constructor(parseException) {
        this.sheet = parseException.Sheet;
        this.formatType = parseException.FormatType;
        this.columnLetter = parseException.ColumnLetter;
        this.columnHeaderName = parseException.ColumnHeader;
        this.row = parseException.Row;
        this.message = parseException.Message;
        this.severity = parseException.Severity === 0 ? 'Error' : 'Warning';
        this.exceptionType = parseException.ExceptionType;
        this.expectedDateType = parseException.ExpectedDateType;
    }

    getDisplayMessage() {
        let sheetMsg = '';
        if (this.sheet != null) {
            sheetMsg = `on sheet "${this.sheet}"`;
        }

        let excelLocation = '';
        if (this.columnHeaderName != null) {
            excelLocation = `at column named "${this.columnHeaderName}"`;
        }
        else if (this.row > 0 && this.columnLetter != null) {
            excelLocation = `at cell "${this.columnLetter}${this.row}"`;
        } else if (this.row > 0) {
            excelLocation = `at row "${this.row}"`;
        } else if (this.columnLetter != null) {
            excelLocation = `at column "${this.columnLetter}"`;
        }
        let expectedDataType = '';
        if (this.expectedDateType != null) {
            expectedDataType = `Expected "${this.expectedDateType}". `;
        }

        //TODO define enums for these in the jsx

        let exceptionTypeMessage = '';
        //UnExpectedDataType = 0
        if (this.exceptionType === 0) {
            exceptionTypeMessage = 'Unexpected data type'
        }
        //DuplicateData = 1
        if (this.exceptionType === 1) {
            exceptionTypeMessage = 'Duplicate data found'
        }
        //DuplicateKey = 2
        if (this.exceptionType === 2) {
            exceptionTypeMessage = 'Duplicate key found'
        }
        //Generic = 3
        if (this.exceptionType === 3) {
            exceptionTypeMessage = this.message;
            this.message = null; //null this out so we swap the message for generic types
        }
        //InvalidData = 4
        if (this.exceptionType === 4) {
            exceptionTypeMessage = 'Invalid data found'
        }
        //MaxLength = 5
        if (this.exceptionType === 5) {
            exceptionTypeMessage = 'Exceeded character limit'
        }
        //MissingData = 6
        if (this.exceptionType === 6) {
            exceptionTypeMessage = 'Missing data'
        }
        //NoFileFound = 7
        if (this.exceptionType === 7) {
            exceptionTypeMessage = 'No file found'
        }
        //NoValidDataToSave = 8
        if (this.exceptionType === 8) {
            exceptionTypeMessage = 'No valid data found'
        }
        //OptionalFieldMissing = 9
        if (this.exceptionType === 9) {
            exceptionTypeMessage = 'Optional field not found'
        }
        //RequiredFieldMissing = 10
        if (this.exceptionType === 10) {
            exceptionTypeMessage = 'Required field not found'
        }
        //SheetMissingError = 11
        if (this.exceptionType === 11) {
            exceptionTypeMessage = 'Sheet not found'
        }
        //WrongFilePassword = 12
        if (this.exceptionType === 12) {
            exceptionTypeMessage = 'Invalid password provided'
        }
        exceptionTypeMessage = exceptionTypeMessage + ' ';


        let location = '';
        if (sheetMsg != '' && excelLocation != '') {
            location = `${sheetMsg}, ${excelLocation}. `;
        }
        else if (sheetMsg != '') {
            location = `${sheetMsg}. `; 
        }
        else if (excelLocation != '') {
            location = `${excelLocation}. `; 
        }

        let msg = ''
        if (this.message != null && this.message !== '') {
            msg = this.message;
        }

        //TODO; would be useful to split these method out in order to bold them and add emphasis later... 
        let finalMessage = `${exceptionTypeMessage}${location}${expectedDataType}${msg}`;

        return finalMessage;
    }
}
Exception.propTypes = {
    sheet: PropTypes.string,
    formatType: PropTypes.string,
    columnLetter: PropTypes.string,
    columnHeaderName: PropTypes.string,
    row: PropTypes.int,
    message: PropTypes.string,
    severity: PropTypes.int,
    exceptionType: PropTypes.string,
    expectedDateType: PropTypes.string,
};

const ParseException = (props) => {
    const exception = new Exception(props.exception);
    console.log(exception);
    //severity 0: Error, severity 1: Warning
    const bootsrapClass = exception.severity === 'Error' ? 'danger' : 'warning';
    const displayMessge = exception.getDisplayMessage();
    return (
        <li className={`list-group-item list-group-item-${bootsrapClass}`}>
            <b>{exception.severity}.</b> {displayMessge}
        </li>
    );
};
ParseException.propTypes = {
    exception: PropTypes.object,
};
