import React from 'react';
import ParseExceptions from './ParseExceptions';
import PropTypes from 'prop-types';

const ParseExceptionsModal = (props) => {
    return (
        <>
            <div
                id="parse-exceptions-modal"
                className="modal fade"
                tabIndex="-1"
                role="dialog"
                aria-labelledby="parse-exceptions-modal"
                aria-hidden="true">
                <div className="modal-dialog" role="document">
                    <div className="modal-content">
                        <div className="modal-header">
                            <h5 className="modal-title">{ props.title}</h5>
                            <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                                <span aria-hidden="true">&times;</span>
                            </button>
                        </div>
                        <div className="modal-body">
                            <ParseExceptions json={props.json} />
                        </div>
                        <div className="modal-footer">
                            <button
                                data-dismiss="modal"
                                id="Cancel"
                                type="button"
                                className="btn btn-outline-secondary">
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
