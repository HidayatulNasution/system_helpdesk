<div class="w-full mt-7">
    <p class="text-xl pb-3 flex items-center">
        <i class="fas fa-list mr-3"></i> Done Tiket
    </p>
    <div class="w-full mt-2 mb-2">
        <div class="flex space-x-4">
            <select id="filter-month" class="border border-gray-400 rounded px-4 py-2">
                <option value="">All Months</option>
                <option value="1">January</option>
                <option value="2">February</option>
                <option value="3">March</option>
                <option value="4">April</option>
                <option value="5">May</option>
                <option value="6">June</option>
                <option value="7">July</option>
                <option value="8">August</option>
                <option value="9">September</option>
                <option value="10">October</option>
                <option value="11">November</option>
                <option value="12">December</option>
            </select>
            <select id="filter-year" class="border border-gray-400 rounded px-4 py-2">
                <option value="">All Years</option>
                @for ($year = 2024; $year <= date('Y') + 6; $year++)
                    <option value="{{ $year }}">{{ $year }}</option>
                @endfor
            </select>
            <button id="filter-button" class="bg-blue-400 btn-warning text-white px-4 py-2 rounded mr-3">🔍
                Filter</button>
            {{-- <a href="javascript:void(0)" id="download-excel"
                class="bg-green-500 text-primary px-4 py-2 rounded">Download
                Excel</a> --}}
            <a href="javascript:void(0)" id="download-pdf">
                <img src="/img/pdfi.png" alt="pdf" height="43px" width="43px">
                Report Tiket
            </a>

        </div>
    </div>
    {{-- <a href="javascript:void(0)" class="btn btn-info ml-3" id="create-new-product">Add New</a> --}}

    <div class="bg-white overflow-auto shadow-md rounded-lg">
        <table class="table-auto w-full border-collapse" id="done_tiket">
            <thead class="bg-gray-800 text-white">
                <tr>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm">No</th>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm" style="min-width: 100px;">Tanggal
                        Entry</th>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm">Username</th>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm">Bidang System</th>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm">Kategori</th>
                    <th class="text-left py-3 px-2 uppercase font-semibold text-sm">Prioritas</th>
                    <th class="text-left py-3 px-2 uppercase font-semibold text-sm" style="min-width:100px;">Status</th>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm" style="min-width:100px;">Tanggal
                        Selesai</th>
                    <th class="text-left py-4 px-2 uppercase font-semibold text-sm" style="min-width:100px;">Action</th>
                </tr>
            </thead>

        </table>
    </div>


</div>

@push('js')
    <script>
        document.getElementById('download-pdf').addEventListener('click', function() {
            const month = document.getElementById('filter-month').value;
            const year = document.getElementById('filter-year').value;

            // Redirect to the PDF download route with query parameters for month and year
            const url = `{{ route('generatePDF') }}?month=${month}&year=${year}`;
            window.location.href = url;
        });
    </script>

    <script>
        // Define `table` in the global scope
        var table;
        var SITEURL = 'http://127.0.0.1:8000/';
        console.log(SITEURL);
        $(document).ready(function() {
            $.ajaxSetup({
                headers: {
                    'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                }
            });
            table = $('#done_tiket').DataTable({
                processing: true,
                serverSide: true,
                ajax: {
                    url: SITEURL + "admin/Done",
                    type: 'GET',
                    data: function(d) {
                        d.month = $('#filter-month').val();
                        d.year = $('#filter-year').val();
                    }
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
                                            `${hours}h ${minutes}m ${seconds}s`;
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
                        data: 'updated_at',
                        name: 'updated_at',
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
                        data: 'action',
                        name: 'action',
                        orderable: false
                    },
                ],
                order: [
                    [0, 'desc']
                ]
            });

            // Click event for the filter button, calling table.draw() if table is defined
            $('#filter-button').click(function() {
                if (table) {
                    table.draw();
                } else {
                    console.error("Table is not defined or initialized.");
                }
            });

            $('#create-new-tiket').click(function() {
                $('#btn-save').val("create-tiket");
                $('#tiket_id').val('');
                $('#tiketForm').trigger("reset");
                $('#tiketCrudModal').html("Add New tiket");
                $('#ajax-tiket-modal').modal('show');
                $('#modal-preview').attr('src', 'https://via.placeholder.com/150');
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
            $.get('admin/Detail/' + tiket_id, function(data) {
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
