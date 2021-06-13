import 'bootstrap';

//styles
import '@fortawesome/fontawesome-svg-core';
import '@fortawesome/free-solid-svg-icons';
import '@fortawesome/free-brands-svg-icons';
import '@fortawesome/react-fontawesome';
import '@fortawesome/free-regular-svg-icons';
import '../Styles/scss/main.scss';

import React from 'react';
import ReactDOM from 'react-dom';
import UploadForm from './UploadForm';
import ParseExceptionsModal from './ParseExceptionsModal';

/*global model*/
/*eslint no-undef: "error"*/

ReactDOM.render(
    <UploadForm
        title="Scan Table Parser"
        export="ExportTableSample"
        import="ScanForColumnsAndParseTable"
        sheetName="Sample Sheet"
        controller="Home"
        jsonParseValue={model.ScanForColumnsAndParseTable}
    />,
    document.getElementById('main')
);
ReactDOM.render(
    <ParseExceptionsModal exceptions={model.Exceptions} title="Import messages" />,
    document.getElementById('modal')
);
