﻿import React, { useState, useEffect } from 'react';
import PropTypes from 'prop-types';
//import Skeleton from 'react-loading-skeleton';

//styles
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faChevronDown, faChevronLeft } from '@fortawesome/free-solid-svg-icons';

function syntaxHighlight(obj) {
    if (obj !== undefined) {
        let json = JSON.stringify(obj, undefined, 4);
        json = json.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return json.replace(
            /("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+/-]?\d+)?)/g,
            function (match) {
                var cls = 'number';
                if (/^"/.test(match)) {
                    if (/:$/.test(match)) {
                        cls = 'key';
                    } else {
                        cls = 'string';
                    }
                } else if (/true|false/.test(match)) {
                    cls = 'boolean';
                } else if (/null/.test(match)) {
                    cls = 'null';
                }
                return `<span class="${cls}">${match}</span>`;
            }
        );
    }
    return `<span></span>`;
}

const JSONOutput = (props) => {
    const [isExpanded, setIsExpanded] = useState(false);
    const [output, setOutput] = useState("");
    //TODO USE REF
    const collapseId = `${props.id}-output`;
    const cheveronIconId = `${props.id}-icon`;

    //Handle collapse click
    function handleCollapseClick() {
        setIsExpanded(!isExpanded);
    }

    useEffect(() => {
        setOutput(syntaxHighlight(props.value));
    }, []);

    //TODO: fix this colappse funciotn 
    useEffect(() => {
        if (!isExpanded) {
            setTimeout(() => {
                //(`#${collapseId}`).removeClass('show');
            }, 351);
        }
    }, [isExpanded]);

    return (
        <div className="output-content">
            <h5 className="mb-0">
                JSON output
                <a
                    id={props.id}
                    className="font-weight-bold pull-right"
                    data-toggle="collapse"
                    href={`#${collapseId}`}
                    aria-expanded={isExpanded}
                    aria-controls={collapseId}
                    title={`${isExpanded ? 'Collapse' : 'Expand'}`}
                    onClick={handleCollapseClick}>
                    <div className="header-chevron" id={cheveronIconId}>
                        {isExpanded ? (
                            <FontAwesomeIcon icon={faChevronDown} />
                        ) : (
                            <FontAwesomeIcon icon={faChevronLeft} />
                        )}
                    </div>
                </a>
            </h5>
            <div className={`row mb-2 pt-2 collapse`} id={collapseId}>
                <div className="col-12 px-0">
                    <div className="json-output" id={`${props.id}-json`}>
                        <pre className="code-wrap" id={`${props.id}-json-txt`}>
                            {output}
                        </pre>
                    </div>
                </div>
            </div>
        </div>
    );
};
JSONOutput.propTypes = {
    id: PropTypes.string,
    value: PropTypes.object,
};

export default JSONOutput;
