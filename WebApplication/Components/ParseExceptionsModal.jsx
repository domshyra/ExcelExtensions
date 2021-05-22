import { Modal } from 'bootstrap';
import React, { useState, useEffect, useRef } from 'react';
import ParseExceptions from './ParseExceptions';
import PropTypes from 'prop-types';

const ParseExceptionsModal = (props) => {
    const [modal, setModal] = useState(null);
    const parseExceptionModal = useRef();


    //asked an issue https://stackoverflow.com/questions/67653772/issue-showing-modal-with-react-hooks-on-bootstrap-5

    //https://dev.to/tefoh/use-bootstrap-5-in-react-2l4i
    useEffect(() => {
        setModal(new Modal(parseExceptionModal.current), {
            keyboard: false,
        });
        console.log(parseExceptionModal.current);
        console.log(parseExceptionModal);
        if (props.json.length > 0) {
            modal.show();
        }
    }, []);

    return (
        <>
            <div
                className="modal fade"
                tabIndex="-1"
                role="dialog"
                ref={parseExceptionModal}
                aria-labelledby="parseExceptionModal"
                aria-hidden="true">
                <div className="modal-dialog" role="document">
                    <div className="modal-content">
                        <div className="modal-header">
                            <h5 className="modal-title">{props.title}</h5>
                            <button type="button" className="btn-close" onClick={() => modal.hide()} aria-label="Close">
                                {/*<span aria-hidden="true">&times;</span>*/}
                            </button>
                        </div>
                        <div className="modal-body">
                            <ParseExceptions json={props.json} />
                        </div>
                        <div className="modal-footer">
                            <button onClick={() => modal.hide()} type="button" className="btn btn-outline-secondary">
                                Dismiss
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </>
    );
};
ParseExceptionsModal.propTypes = {
    json: PropTypes.array,
    title: PropTypes.string,
};
export default ParseExceptionsModal;
