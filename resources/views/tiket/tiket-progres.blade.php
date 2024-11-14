<div class="w-full mt-7">
    <p class="text-xl pb-3 flex items-center">
        <i class="fas fa-list mr-3"></i> Progress Tiket
    </p>
    <a href="javascript:void(0)" class="btn btn-primary ml-3" id="create-new-tiket">➕ Add New</a>

    <div class="bg-white overflow-auto shadow-md rounded-lg">
        <table class="table-auto w-full border-collapse" id="data_tiket">
            <thead class="bg-gray-800 text-white">
                <tr>
                    <th class="text-left py-4 px-3 uppercase font-bold text-sm">No</th>
                    <th class="text-left py-4 px-4 uppercase font-bold text-sm" style="min-width: 100px;">Tanggal
                        Entry</th>
                    <th class="text-left py-4 px-4 uppercase font-semibold text-sm">Username</th>
                    <th class="text-left py-4 px-4 uppercase font-semibold text-sm">Bidang System</th>
                    <th class="text-left py-4 px-4 uppercase font-semibold text-sm">Kategori</th>
                    <th class="text-left py-4 px-4 uppercase font-semibold text-sm">Prioritas</th>
                    <th class="text-left py-4 px-5 uppercase font-semibold text-sm" style="min-width:100px;">Status</th>
                    <th class="text-left py-4 px-5 uppercase font-semibold text-sm" style="min-width:100px;">Action</th>
                </tr>
            </thead>

        </table>
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
            $('#data_tiket').DataTable({
                processing: true,
                serverSide: true,
                ajax: {
                    url: SITEURL + "dashboard",
                    type: 'GET',
                },
                columns: [{
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
                        data: 'bidang_system',
                        name: 'bidang_system'
                    },
                    {
                        data: 'kategori',
                        name: 'kategori'
                    },
                    {
                        data: 'prioritas',
                        name: 'prioritas',
                        render: function(data, type, row) {
                            return data == 1 ? 'URGENT' : 'BIASA';
                        }
                    },
                    {
                        data: 'status',
                        name: 'status',
                        render: function(data, type, row) {
                            let statusText = data == 0 ? 'On Progress' : 'DONE';

                            // Calculate time remaining
                            if (data == 0 && row.created_at) {
                                const createdAt = new Date(row.created_at);
                                const deadline = new Date(createdAt.getTime() + 24 * 60 * 60 *
                                    1000); // 24 hours later

                                // Generate a unique ID for each row
                                const uniqueId = 'time-remaining-' + row
                                    .id; // Assuming `row.id` is unique for each row
                                statusText += ` (Waktu Tersisa <span id="${uniqueId}"></span>)`;

                                // Update the time remaining every second
                                setTimeout(function updateCountdown() {
                                    const now = new Date();
                                    const timeRemaining = deadline - now;

                                    if (timeRemaining > 0) {
                                        // Format time remaining
                                        const hours = Math.floor(timeRemaining / (1000 *
                                            60 * 60));
                                        const minutes = Math.floor((timeRemaining % (1000 *
                                            60 * 60)) / (1000 * 60));
                                        const seconds = Math.floor((timeRemaining % (1000 *
                                            60)) / 1000);
                                        document.getElementById(uniqueId).innerText =
                                            `${hours}: ${minutes}: ${seconds} `;
                                        setTimeout(updateCountdown,
                                            1000); // Repeat every second
                                    } else {
                                        document.getElementById(uniqueId).innerText =
                                            'Time Expired';
                                    }
                                }, 0);
                            }

                            return statusText;
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

            $('#create-new-tiket').click(function() {
                $('#btn-save').val("create-tiket");
                $('#tiket_id').val('');
                $('#tiketForm').trigger("reset");
                $('#tiketCrudModal').html("New Tiket");
                $('#ajax-tiket-modal').modal('show');
                $('#modal-preview').attr('src', 'https://via.placeholder.com/150');
            });

            $('#btn-cancel').click(function() {
                $('#ajax-tiket-modal').modal('hide');
            })

            $('body').on('click', '.edit-tiket', function() {
                var tiket_id = $(this).data('id');
                console.log(tiket_id);
                $.get('dashboard/Edit/' + tiket_id, function(data) {
                    $('#title-error').hide();
                    $('#tiket_category-error').hide();
                    $('#category-error').hide();
                    $('#tiketCrudModal').html("Edit Tiket");
                    $('#btn-save').val("edit-tiket");
                    $('#ajax-tiket-modal').modal('show');
                    $('#tiket_id').val(data.id);
                    $('#bidang_system').val(data.bidang_system);
                    $('#kategori').val(data.kategori);
                    $('#problem').val(data.problem);
                    $('#status').val(data.status);
                    //$('#result').val(data.result);
                    $('#prioritas').val(data.prioritas);
                    $('#modal-preview').attr('alt', 'No image available');
                    if (data.image) {
                        $('#modal-preview').attr('src', SITEURL + 'public/product/' + data.image);
                        $('#hidden_image').attr('src', SITEURL + 'public/product/' + data.image);
                    }
                })
            });

            $('body').on('click', '#delete-tiket', function() {
                var tiket_id = $(this).data("id");
                if (confirm("Are You sure want to delete !")) {
                    $.ajax({
                        type: "GET",
                        url: SITEURL + "dashboard/Delete/" + tiket_id,
                        success: function(data) {
                            var oTable = $('#data_tiket').dataTable();
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

        $('body').on('submit', '#tiketForm', function(e) {
            e.preventDefault();
            var actionType = $('#btn-save').val();
            $('#btn-save').html('Sending..');
            var formData = new FormData(this);
            $.ajax({
                type: 'POST',
                url: SITEURL + "dashboard/Store",
                data: formData,
                cache: false,
                contentType: false,
                processData: false,
                success: (data) => {
                    console.log('Data Saved: ', data);
                    Swal.fire({
                        title: "Good job!",
                        text: "Ticket created and notification sent successfully!",
                        icon: "success"

                    });

                    $('#tiketForm').trigger("reset");
                    $('#ajax-tiket-modal').modal('hide');
                    $('#btn-save').html('Save Changes');
                    var oTable = $('#data_tiket').dataTable();
                    oTable.fnDraw(false);
                    location.reload();
                },
                error: function(data) {
                    console.log('Error:', data);
                    $('#btn-save').html('Save Changes');
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

        // Detail button click
        $('body').on('click', '.detail-tiket', function() {
            var tiket_id = $(this).data('id');
            //console.log(data);
            $.get('dashboard/Detail/' + tiket_id, function(data) {
                if (data) {
                    // Fill the modal with data
                    $('#detail-bidang_system').val(data.bidang_system);
                    $('#detail-kategori').val(data.kategori);
                    $('#detail-problem').val(data.problem);
                    $('#detail-result').val(data.result);
                    $('#detail-prioritas').val(data.prioritas == 0 ? 'BIASA' : 'URGENT');
                    $('#detail-status').val(data.status == 0 ? 'On Progress' : 'DONE');

                    if (data.image) {
                        const imageUrl = SITEURL + 'public/product/' + data.image;
                        $('#modal-preview').attr('src', imageUrl);
                        $('#hidden_image').attr('src', imageUrl);
                        $('#image-url').val(imageUrl); // Set the hidden image URL for viewing
                    }

                    // Show the modal
                    $('#detailTiketModal').modal('show');
                } else {
                    alert('Data tidak ditemukan.');
                }
            }).fail(function() {
                alert('Gagal mengambil data.');
            });
        });
        // Close modal
        $('#btn-close').click(function() {
            $('#detailTiketModal').modal('hide');
        });

        // JavaScript to open a new page to display the image
        document.getElementById('viewImageBtn').addEventListener('click', function() {
            const imageUrl = document.getElementById('image-url').value;
            if (imageUrl) {
                window.open(imageUrl, '_blank');
            } else {
                Swal.fire({
                    icon: "error",
                    title: "Empty",
                    text: "No image available to view!",
                });
            }
        });
    </script>
@endpush
