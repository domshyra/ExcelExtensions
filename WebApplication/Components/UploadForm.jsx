import React, { useState, useEffect } from 'react';
import PropTypes from 'prop-types';

const UploadForm = (props) => {
    useEffect(() => {
        //later
    }, []);

    return (
        <>
            <h2 className="text-primary">{props.title}</h2>
            <div className="row mb-2">
                <div className="col-12">
                    <div className="bg-light px-3 pt-2 pb-3 border rounded">
                        <form
                            action={props.action}
                            controller="Home"
                            encType="multipart/form-data"
                            method="post"
                            className="form-horizontal needs-validation"
                            noValidate>
                            <div className="form-group">
                                <label htmlFor="formFile" className="form-label">
                                    Table file
                                </label>
                                <input className="form-control form-control-sm" type="file" id="formFile" />
                                <small className="form-text text-muted">
                                    This requires an excel sheet a sheet named {props.sheetName}.
                                </small>
                            </div>
                            <button className="btn btn-outline-primary" type="submit">
                                Import
                            </button>
                        </form>
                        {/*TODO: add back react component*/}
                    </div>
                </div>
            </div>
        </>
    );
};
UploadForm.propTypes = {
    title: PropTypes.string,
    action: PropTypes.string,
    sheetName: PropTypes.string,
};

export default UploadForm;
