import React from 'react';
import bsCustomFileInput from 'bs-custom-file-input';


function UploadForm() {

    useEffect(() => {
        bsCustomFileInput.init()
    }, []);

    return (
        <>
            <h2 className="text-primary">{props.FormTitle}</h2>
            <div className="row mb-2">
                <div className="col-12">
                    <div className="bg-light px-3 pt-2 pb-3 border rounded">
                        <form asp-action="ImportTable" asp-controller="Home" enctype="multipart/form-data" method="post" className="form-horizontal needs-validation" novalidate>
                            <div className="form-group">
                                <label>Table file</label>
                                <div className="custom-file">
                                    <input type="file" className="custom-file-input" name="file" accept=".xls,.xlsx" required />
                                    <label className="custom-file-label" for="customFile">Choose file</label>
                                </div>
                                    <small className="form-text text-muted">
                                        This requires an excel sheet a sheet named { props.SheetName }.
                                    </small>
                                </div>
                                <button className="btn btn-primary" type="submit">Import</button>
                        </form>
                        {/*TODO: add back react component*/}
                    </div>
                </div>
            </div>
        </>
    );
}
