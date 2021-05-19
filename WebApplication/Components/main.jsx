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

ReactDOM.render(<UploadForm title="Test" action="test" sheetName="Sheet1" />, document.getElementById('main'));
