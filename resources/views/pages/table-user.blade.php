<div class="row mt-4 mx-4">
    <div class="col-12">
        <div class="alert alert-light" role="alert">
            Manage User</strong>. Check it
            <strong>
                <a href="https://www.creative-tim.com/product/argon-dashboard-pro-laravel" target="_blank">
                    Manage user
                </a>
            </strong>
        </div>
        <div class="card mb-4">
            <button
                class="w-20 bg-skyblue cta-btn font-semibold py-2 rounded-br-lg rounded-bl-lg rounded-tr-lg shadow-lg hover:shadow-xl hover:bg-gray-300 hover:scale-105 transform transition-transform duration-200 flex items-center justify-center">
                <i class="fas fa-user-plus "></i> <a href="javascript:void(0)" id="create-new-user">New User</a>
            </button>
            <div class="card-body px-0 pt-0 pb-2 mt-4">
                <div class="table-responsive p-0">
                    <table class="table align-items-center mb-0" id="laravel_11_datatable2">
                        <thead>
                            <tr>
                                <th class="text-uppercase text-secondary text-xxs font-weight-bolder opacity-7">No
                                </th>
                                <th class="text-uppercase text-secondary text-xxs font-weight-bolder opacity-7 ps-2">
                                    Create Date
                                </th>
                                <th
                                    class="text-center text-uppercase text-secondary text-xxs font-weight-bolder opacity-7">
                                    Username</th>
                                <th
                                    class="text-center text-uppercase text-secondary text-xxs font-weight-bolder opacity-7">
                                    Email</th>
                                <th
                                    class="text-center text-uppercase text-secondary text-xxs font-weight-bolder opacity-7">
                                    Status</th>
                                <th
                                    class="text-center text-uppercase text-secondary text-xxs font-weight-bolder opacity-7">
                                    Action</th>
                            </tr>
                        </thead>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

@push('js')
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    <script>
        var SITEURL = 'http://127.0.0.1:8000/';
        console.log(SITEURL);
        $(document).ready(function() {
            $.ajaxSetup({
                headers: {
                    'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                }
            });
            $('#laravel_11_datatable2').DataTable({
                processing: true,
                serverSide: true,
                ajax: {
                    url: SITEURL + "user",
                    type: 'GET',
                },
                columns: [

                    {
                        data: 'DT_RowIndex',
                        name: 'DT_RowIndex',
                        orderable: false,
                        searchable: false
                    },
                    {
                        data: 'created_at',
                        name: 'created_at',
                        render: function(data, type, row) {
                            if (data) {
                                const date = new Date(data);
                                const formattedDate =
                                    ('0' + (date.getMonth() + 1)).slice(-2) + '-' +
                                    ('0' + date.getDate()).slice(-2) + '-' +
                                    date.getFullYear() + ' ' +
                                    ('0' + date.getHours()).slice(-2) + ':' +
                                    ('0' + date.getMinutes()).slice(-2) + ':' +
                                    ('0' + date.getSeconds()).slice(-2);
                                return formattedDate;
                            }
                            return '';
                        }
                    },
                    {
                        data: 'username',
                        name: 'username'
                    },
                    {
                        data: 'email',
                        name: 'email'
                    },

                    {
                        data: 'status',
                        name: 'status',
                        render: function(data, type, row) {
                            return data == 1 ? 'ADMIN' : 'USER';
                        }
                    },
                    {
                        data: 'action',
                        name: 'action',
                        orderable: false
                    },
                ],
                order: [
                    [0, 'desc']
                ]
            });

            $('#create-new-user').click(function() {
                $('#btn-save').val("create-user");
                $('#user_id').val('');
                $('#userForm').trigger("reset");
                $('#userCrudModal').html("Add New user");
                $('#ajax-user-modal').modal('show');
                $('#modal-preview').attr('src', 'https://via.placeholder.com/150');
            });

            $('#btn-cancel').click(function() {
                $('#ajax-user-modal').modal('hide');
            })

            $('body').on('click', '.edit-user', function() {
                var user_id = $(this).data('id');
                console.log(user_id);
                $.get('user/Edit/' + user_id, function(data) {
                    $('#userCrudModal').html("Edit user");
                    $('#btn-save').val("edit-user");
                    $('#ajax-user-modal').modal('show');
                    $('#user_id').val(data.id);
                    $('#created_at').val(data.created_at);
                    $('#username').val(data.username);
                    $('#email').val(data.email);
                    $('#password').val(data.password);
                    $('#status').val(data.status);
                })
            });


            // Detail button click
            $('body').on('click', '.detail-user', function() {
                var user_id = $(this).data('id');
                //console.log(data);
                $.get('user/Detail/' + user_id, function(data) {
                    if (data) {
                        // Fill the modal with data
                        $('#detail-username').val(data.username);
                        $('#detail-email').val(data.email);
                        //$('#detail-password').val(data.password);
                        $('#detail-status').val(data.status == 0 ? 'ADMIN' : 'USER');

                        // Show the modal
                        $('#detailUserModal').modal('show');
                    } else {
                        alert('Data tidak ditemukan.');
                    }
                }).fail(function() {
                    alert('Gagal mengambil data.');
                });
            });
            // Close modal
            $('#btn-close').click(function() {
                $('#detailUserModal').modal('hide');
            });

            $('body').on('click', '#delete-user', function() {
                var user_id = $(this).data("id");

                Swal.fire({
                    title: "Apa Anda Yakin ?",
                    text: "Anda tidak akan dapat mengembalikan data ini !",
                    icon: "warning",
                    showCancelButton: true,
                    confirmButtonColor: "#3085d6",
                    cancelButtonColor: "#d33",
                    confirmButtonText: "Yes, delete it!"
                }).then((result) => {
                    if (result.isConfirmed) {
                        $.ajax({
                            type: "GET",
                            url: SITEURL + "user/Delete/" + user_id,
                            success: function(data) {
                                Swal.fire("Deleted!", "Data Anda telah dihapus.",
                                    "success");

                                var oTable = $('#laravel_11_datatable2').dataTable();
                                oTable.fnDraw(false);
                                location.reload();
                            },
                            error: function(data) {
                                console.log('Error:', data);
                            }
                        });
                    }
                });
            });


        });

        $('body').on('submit', '#userForm', function(e) {
            e.preventDefault();

            Swal.fire({
                title: "Do you want to save this data?",
                showDenyButton: true,
                showCancelButton: true,
                confirmButtonText: "Save",
                denyButtonText: `Don't save`
            }).then((result) => {
                if (result.isConfirmed) {
                    Swal.fire("successfully Add User !", "", "success");

                    var actionType = $('#btn-save').val();
                    $('#btn-save').html('Sending..');
                    var formData = new FormData(this);

                    $.ajax({
                        type: 'POST',
                        url: SITEURL + "user/Store",
                        data: formData,
                        cache: false,
                        contentType: false,
                        processData: false,
                        success: function(data) {
                            console.log(data);

                            $('#userForm').trigger("reset");
                            $('#ajax-user-modal').modal('hide');
                            $('#btn-save').html('Save Changes');
                            var oTable = $('#laravel_11_datatable2').dataTable();
                            oTable.fnDraw(false);
                            location.reload();
                        },
                        error: function(data) {
                            console.log('Error:', data);
                            $('#btn-save').html('Save Changes');
                        }
                    });
                } else if (result.isDenied) {
                    Swal.fire("Data are not saved", "", "info");
                }
            });
        });



        function readURL(input, id) {
            id = id || '#modal-preview';
            if (input.files && input.files[0]) {
                var reader = new FileReader();
                reader.onload = function(e) {
                    $(id).attr('src', e.target.result);
                };
                reader.readAsDataURL(input.files[0]);
                $('#modal-preview').removeClass('hidden');
                $('#start').hide();
            }
        }
    </script>
@endpush
