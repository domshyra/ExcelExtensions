import React from 'react';

function UploadForm() {
    return (
        <>
            <h2 className="text-primary">{props.FormTitle}</h2>
            <div className="row mb-2">
                <div className="col-12">
                    <div className="bg-light px-3 pt-2 pb-3 border rounded">
                        <form asp-action="SearchAndImportTable" asp-controller="ExcelExtensions" enctype="multipart/form-data" method="post" className="form-horizontal needs-validation" novalidate>
                            <div className="form-group">
                                <label>Table line items</label>
                                <div className="custom-file">
                                    <input type="file" className="custom-file-input" name="file" accept=".xls,.xlsx" required>
                                        <label className="custom-file-label" for="customFile">Choose file</label>
                                        </div>
                                    <small className="form-text text-muted">
                                        This requires an excel sheet a sheet named "Sheet1".
                                        </small>
                                </div>

                                <button className="btn btn-primary" type="submit">Import</button>
                            </div>

                        </form>
                        {/*TODO:Make this output into a react componet*/}
                        <div className="playground-output-content">
                            <h5 className="mb-0">
                                JSON output
                                <a id="excel-search-table-json-output-toggler" className="font-weight-bold pull-right" data-toggle="collapse" href="#search-table-json-output" aria-expanded="true" aria-controls="json-output">
                                    <i className="fa fa-chevron-down cb-header-chevron"></i>
                                </a>
                            </h5>
                            <div className="row collapse show mb-2 pt-2" id="search-table-json-output">
                                <div className="col-12">
                                    <div className="playground-json-output" id="search-table-json">
                                        <pre className="playground-code-wrap" id="search-table-json-txt"></pre>
                                    </div>
                                    <small className="form-text text-muted">
                                        Json format is "Row #": {rowModel}
                                    </small>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </>
    );
}
