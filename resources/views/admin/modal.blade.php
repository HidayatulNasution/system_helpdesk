{{-- Modal Insert --}}
<div class="modal fade" id="ajax-tiket-modal2" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header" style="background-color: rgba(0, 202, 238)">
                <h4 class="modal-title" id="tiketCrudModal2"></h4>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form id="tiketForm2" name="tiketForm2" class="form-horizontal" enctype="multipart/form-data">
                    <div class="row">
                        <div class="col-xl-14 col-lg-14">
                            <div class="card">
                                <div class="card-body">
                                    <div class="row">
                                        <input type="hidden" name="tiket_id" id="tiket_id">
                                        <input type="hidden" name="created_at" id="created_at"
                                            value="<?= date('Y-m-d') ?>">
                                        <div class="form-group">
                                            <label for="bidang_system" class="col-sm-4 control-label">Bidang
                                                System</label>
                                            <div class="col-sm-12">
                                                <select class="form-control" name="bidang_system" id="bidang_system"
                                                    required>
                                                    <option value=""> -- Pilih Bidang System -- </option>
                                                    <option value="CRM">CRM</option>
                                                    <option value="FICO">FICO</option>
                                                    <option value="WORKSHOP">WORKSHOP</option>
                                                </select>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label for="kategori" class="col-sm-4 control-label">Kategori</label>
                                            <div class="col-sm-12">
                                                <select class="form-control" name="kategori" id="kategori" required>
                                                    <option value=""> -- Pilih Kategori -- </option>
                                                    <option value="CRM">CRM</option>
                                                    <option value="FICO">FICO</option>
                                                    <option value="WORKSHOP">WORKSHOP</option>
                                                </select>
                                            </div>
                                        </div>

                                        <div class="form-group">
                                            <label for="status" class="col-sm-4 control-label">Status</label>
                                            <select class="form-control" name="status" id="status">
                                                <option value="0">On Progress</option>
                                                <option value="1">Done</option>
                                            </select>
                                        </div>
                                        <div class="form-group">
                                            <label for="prioritas" class="col-sm-4 control-label">Prioritas</label>
                                            <select class="form-control" name="prioritas" id="prioritas">
                                                <option value="0">BIASA</option>
                                                <option value="1">URGENT</option>
                                            </select>
                                        </div>
                                        <div class="form-group">
                                            <label for="problem" class="col-sm-4 control-label">Problem</label>
                                            <div class="col-sm-12">
                                                <textarea name="problem" id="problem" cols="45" rows="10" placeholder="Jelaskan Request / Problem di system"
                                                    autocomplete="off"></textarea>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label for="result" class="col-sm-4 control-label">Result</label>
                                            <div class="col-sm-12">
                                                <textarea name="result" id="result" cols="45" rows="10"
                                                    placeholder="Jelaskan Result / Penyelesaian dari problem" autocomplete="off"></textarea>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="col-sm-4 control-label">Lampiran</label>
                                            <div class="col-sm-12">
                                                <input id="image" type="file" name="image" accept="image/*"
                                                    onchange="readURL(this);">
                                                <input type="hidden" name="hidden_image" id="hidden_image">
                                            </div>
                                        </div>
                                        <img id="modal-preview" src="https://via.placeholder.com/100" alt="Preview"
                                            class="form-group hidden" width="100" height="100">
                                        <div class="col-sm-offset-2 col-sm-10">
                                            <button type="submit" class="btn btn-primary" id="btn-save"
                                                value="create">Save
                                                changes</button>
                                            <button type="button" class="btn btn-danger" data-dismiss="modal"
                                                id="btn-cancel">Cancel</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </form>
            </div>
            <div class="modal-footer">
            </div>
        </div>
    </div>
</div>
<!-- Detail Modal -->
<div class="modal fade" id="detailTiketModal" tabindex="-1" role="dialog" aria-labelledby="detailTiketModalLabel"
    aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header" style="background-color: #00e897;">
                <h4 class="modal-title" id="detailTiketModalLabel">Detail Data</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
            </div>
            <div class="modal-body">
                <form action="" id="formDetail">
                    <div class="row">
                        <div class="col-xl-12 col-lg-12">
                            <div class="card">
                                <div class="card-body">
                                    <div class="row">
                                        <div class="form-group col-6">
                                            <label for="bidang_system">bidang_system</label>
                                            <input type="text" class="form-control" id="detail-bidang_system"
                                                name="detail-bidang_system" readonly>
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="kategori">Kategori</label>
                                            <input type="text" class="form-control" id="detail-kategori"
                                                name="detail-kategori" readonly>
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="status">Status</label>
                                            <input type="text" class="form-control" id="detail-status"
                                                name="detail-status" readonly>
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="prioritas">prioritas</label>
                                            <input type="text" class="form-control" id="detail-prioritas"
                                                name="detail-prioritas" readonly>
                                        </div>
                                        <div class="form-group col-12">
                                            <label for="problem">Detail Problem</label>
                                            <textarea class="form-control" name="detail-problem" id="detail-problem" cols="30" rows="15" readonly></textarea>
                                        </div>
                                        <div class="form-group col-12">
                                            <label for="detail-result">Detail Result</label>
                                            <textarea class="form-control" name="detail-result" id="detail-result" cols="30" rows="15" readonly></textarea>
                                        </div>

                                        <!-- Hidden field to store image URL -->
                                        <input type="hidden" id="image-url">

                                        <!-- New View image Button -->
                                        <div class="form-group col-12  mt-3 mb-3   d-flex justify-content-between">

                                            <button type="button" class="btn btn-primary" id="viewImageBtn">
                                                <i class="fas fa-eye"></i> View Image
                                            </button>
                                            <button type="button" class="btn btn-danger" data-dismiss="modal"
                                                id="btn-close"><i class="fas fa-angle-double-left"></i>
                                                Close</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                </form>
            </div>
        </div>
    </div>
</div>
