<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class tiket extends Model
{
    use HasFactory;
    protected $table = 'tikets';
    protected $fillable = [
        'created_at',
        'bidang_system',
        'kategori',
        'status',
        'problem',
        'prioritas',
        'image'
    ];

    // Relasi Many-to-One: Banyak Tiket dimiliki oleh satu User
    public function user()
    {
        return $this->belongsTo(User::class);
    }
}
