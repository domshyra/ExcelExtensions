import React, { useEffect } from 'react';
import PropTypes from 'prop-types';

const UploadForm = (props) => {
    useEffect(() => {
        //later
    }, []);

    return (
        <>
            <h2 className="text-white">{props.title}</h2>
            <div className="row mb-2">
                <div className="col-12">
                    <div className="bg-light px-3 pt-2 pb-3 border rounded">
                        <div className="row">
                            <div className="col-8">
                                <p>Export a sample table with the required columns and sheet name to test the <b>{props.title}</b>.</p>
                            </div>
                            <div className="col-4">
                                <a
                                    role="button"
                                    className="btn btn-sm btn-outline-success float-end"
                                    href={`../${props.controller}/${props.export}`}>
                                    Export example sheet
                        </a>
                            </div>
                        
                        </div>

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
                            <button className="btn btn-primary" type="submit">
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
    import: PropTypes.string,
    export: PropTypes.string,
    controller: PropTypes.string,
    sheetName: PropTypes.string,
};

export default UploadForm;
