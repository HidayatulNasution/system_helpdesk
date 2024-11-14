<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class admin extends Model
{
    use HasFactory;
    protected $table = 'tikets';
    protected $fillable = [
        'created_at',
        'bidang_system',
        'kategori',
        'status',
        'problem',
        'result',
        'prioritas',
        'image'
    ];

    public function user()
    {
        return $this->belongsTo(User::class);
    }
}
