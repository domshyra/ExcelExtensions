import { Modal } from 'bootstrap';
import React, { useState, useEffect, useRef } from 'react';
import ParseExceptions from './ParseExceptions';
import PropTypes from 'prop-types';

const ParseExceptionsModal = (props) => {
    const [modal, setModal] = useState(null);
    const parseExceptionModal = useRef();

    useEffect(() => {
        const modal = new Modal(parseExceptionModal.current, { keyboard: false });
        setModal(modal);
        if (props.exceptions.length > 0) {
            modal.show();
            console.log(props.exceptions);
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
                <div className="modal-dialog modal-lg" role="document">
                    <div className="modal-content">
                        <div className="modal-header">
                            <h5 className="modal-title">{props.title}</h5>
                            <button type="button" className="btn-close" onClick={() => modal.hide()} aria-label="Close">
                                {/*<span aria-hidden="true">&times;</span>*/}
                            </button>
                        </div>
                        <div className="modal-body">
                            <ParseExceptions json={props.exceptions} />
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
    exceptions: PropTypes.array,
    title: PropTypes.string,
};
export default ParseExceptionsModal;
