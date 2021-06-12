import React, { useEffect } from 'react';
import PropTypes from 'prop-types';

const UploadForm = (props) => {
    useEffect(() => {
        //later
    }, []);

    //function handleSubmit(event) {

    //};

    return (
        <>
            <h2 className="text-primary">{props.title}</h2>
            <div className="row mb-2">
                <div className="col-12">
                    <div className="bg-light px-3 pt-2 pb-3 border rounded">
                        <form
                            action={`../${props.controller}/${props.import}`}
                            encType="multipart/form-data"
                            method="post"
                            className="form-horizontal needs-validation"
                            noValidate>
                            <div className="form-group mb-3">
                                <label htmlFor="file" className="form-label">
                                    Table file
                                </label>
                                <input className="form-control form-control-sm" type="file" id="file" name="file" />
                                <small className="form-text text-muted">
                                    This requires an excel sheet a sheet named {props.sheetName}.
                                </small>
                            </div>
                            <div className="form-footer">
                                <button className="btn btn-outline-primary pr-2 mr-2" type="submit">
                                    Import
                                </button>
                                <a
                                    role="button"
                                    className="btn btn-outline-success pl-2 ml-2"
                                    href={`../${props.controller}/${props.export}`}>
                                    Export
                                </a>
                            </div>
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
    import: PropTypes.string,
    export: PropTypes.string,
    controller: PropTypes.string,
    sheetName: PropTypes.string,
};

export default UploadForm;
