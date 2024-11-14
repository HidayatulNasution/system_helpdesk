{{-- Modal Update & Insert --}}
<div class="modal fade" id="ajax-user-modal" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header" style="background-color: rgba(0, 202, 238, 0.533);">
                <h4 class="modal-title" id="userCrudModalDetail">User Form</h4>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <form id="userForm" name="userForm" class="form-horizontal" enctype="multipart/form-data">
                    <div class="row">
                        <div class="col-xl-14 col-lg-14">
                            <div class="card">
                                <div class="card-body">
                                    <div class="row">
                                        <input type="hidden" name="user_id" id="user_id">

                                        <div class="form-group">
                                            {{-- <label for="tgl_entry" class="col-sm-4 control-label">Tanggal Entry</label> --}}
                                            <div class="col-sm-12">
                                                <input type="hidden" class="form-control" id="created_at"
                                                    name="created_at" value="<?= date('Y-m-d') ?>">
                                            </div>
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="username">Username</label>
                                            <input type="text" class="form-control" id="username" name="username"
                                                value="" required="" autocomplete="off">
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="email">Email</label>
                                            <input type="text" class="form-control" id="email" name="email"
                                                value="" required="" autocomplete="off">
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="password">Password</label>
                                            <input type="password" class="form-control" id="password" name="password"
                                                autocomplete="off" value="{{ session('plainPassword') }}">
                                        </div>
                                        <div class="form-group">
                                            <input type="checkbox" id="showPassword"
                                                onclick="togglePasswordVisibility()">
                                            <label for="showPassword">Show Password</label>
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="status">Status</label>
                                            <select class="form-control" name="status" id="status" required>
                                                <option value="0">USER</option>
                                                <option value="1">ADMIN</option>
                                            </select>
                                        </div>

                                        <div class="col-sm-offset-2 col-sm-10">
                                            <br>
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
<div class="modal fade" id="detailUserModal" tabindex="-1" role="dialog" aria-labelledby="detailUserModalLabel"
    aria-hidden="true">
    <div class="modal-dialog" role="document">
        <div class="modal-content">
            <div class="modal-header" style="background-color: #00e897;">
                <h4 class="modal-title" id="detailUserModalLabel">Detail Data User</h5>
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
                                            <label for="username">Username</label>
                                            <input type="text" class="form-control" id="detail-username"
                                                name="detail-username" readonly>
                                        </div>
                                        <div class="form-group col-6">
                                            <label for="email">Email</label>
                                            <input type="text" class="form-control" id="detail-email"
                                                name="detail-email" readonly>
                                        </div>
                                        {{-- <div class="form-group col-6">
                                            <label for="password">Detail password</label>
                                            <input type="text" class="form-control" id="detail-password"
                                                name="detail-password" readonly>
                                        </div> --}}
                                        <div class="form-group col-6">
                                            <label for="status">Status</label>
                                            <input type="text" class="form-control" id="detail-status"
                                                name="detail-status" readonly>
                                        </div>

                                        <div class="form-group col-12  mt-3 mb-3   d-flex justify-content-between">
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


@push('js')
    <script>
        function togglePasswordVisibility() {
            var passwordInput = document.getElementById("password");
            var showPasswordCheckbox = document.getElementById("showPassword");

            if (showPasswordCheckbox.checked) {
                passwordInput.type = "text";
            } else {
                passwordInput.type = "password";
            }
        }
    </script>
@endpush
