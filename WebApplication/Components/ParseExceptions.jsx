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
        this.sheet = parseException.sheet;
        this.formatType = parseException.formatType;
        this.columnLetter = parseException.columnLetter;
        this.columnHeaderName = parseException.columnHeader;
        this.row = parseException.row;
        this.message = parseException.message;
        this.severity = parseException.severity;
        this.exceptionType = parseException.exceptionType;
        this.expectedDateType = parseException.expectedDateType;
    }

    getDisplayMessage() {
        let sheetMsg = '';
        if (this.sheet !== '') {
            sheetMsg = `On sheet ${this.sheet}.`;
        }

        let locationMsg = '';
        if (this.row !== '' && this.columnLetter !== '') {
            locationMsg = `At cell ${this.columnLetter}${this.row}.`;
        } else if (this.row !== '') {
            locationMsg = `At row ${this.row}.`;
        } else if (this.columnLetter !== '') {
            locationMsg = `At column ${this.columnLetter}.`;
        } else if (this.columnHeaderName !== '') {
            locationMsg = `At column named ${this.columnHeaderName}.`;
        }

        //TODO maybe split these up to make it easier to write errors
        return `${sheetMsg} ${locationMsg} ${this.expectedDateType}. ${this.message}`;
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

    //severity 0: Error, severity 1: Warning
    const bootsrapClass = exception.severity === 0 ? 'danger' : 'warning';
    const severityMessage = exception.severity === 0 ? 'Error' : 'Warning';
    const displayMessge = exception.getDisplayMessage();
    return (
        <li className={`list-group-item list-group-item-${bootsrapClass}`}>
            <b>{severityMessage}.</b> {displayMessge}
        </li>
    );
};
ParseException.propTypes = {
    exception: PropTypes.object,
};
