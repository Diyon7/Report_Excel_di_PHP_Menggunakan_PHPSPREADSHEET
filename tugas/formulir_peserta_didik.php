<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <!-- <link rel="stylesheet" href="bootstrap.css"> -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css" integrity="sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh" crossorigin="anonymous">
    <script type="text/javascript" src="jquery-3.4.1.min.js"></script>
    <script>
        $(document).ready(function() {
            $("#negara").hide()
            $("#Indonesia").click(function() {
                $("#negara").val(""), $("#negara").hide("fast");
            })
            $("#Asing").click(function() {
                $("#negara").show();
            })
        })
        $(document).ready(function() {
            $("#Tidak").click(function() {
                $("#kps").val(""), $("#kps").attr("disabled", true);
            })
            $("#Ya").click(function() {
                $("#kps").val(""), $("#kps").attr("disabled", false);
            })
        })
    </script>
</head>

<body>
    <?php

    include "koneksi.php";

    date_default_timezone_set('Asia/Jakarta');
    $waktu = date('Y-m-d H:i:s');
    $waktut = date('d-m-Y');
    $jam = date('H:i:s');

    $error_nama = "";
    $error_jk = "";
    $error_nisn = "";
    $error_nik = "";
    $error_telahir = "";
    $error_talahir = "";
    $error_noakta = "";
    $error_agama = "";
    $error_negara = "";
    $error_kebkus = "";
    $error_alamat = "";
    $error_rt = "";
    $error_rw = "";
    $error_dusun = "";
    $error_desa = "";
    $error_kec = "";
    $error_kp = "";
    $error_lintang = "";
    $error_bujur = "";
    $error_temting = "";
    $error_modtrans = "";
    $error_kks = "";
    $error_anakke = "";
    $error_pekps = "";
    $error_kps = "";

    $nama = "";
    $jk = "";
    $nisn = "";
    $nik = "";
    $telahir = "";
    $talahir = "";
    $noakta = "";
    $agama = "";
    $negara = "";
    $na = "";
    $negarar = "";
    $kebkus = "";
    $alamat = "";
    $rt = "";
    $rw = "";
    $dusun = "";
    $desa = "";
    $kec = "";
    $kp = "";
    $lintang = "";
    $bujur = "";
    $temting = "";
    $modtrans = "";
    $kks = "";
    $anakke = "";
    $pekps = "";
    $kps = "";

    if (!empty($_POST["negara"])) {
        $negarar = $_POST["negara"];
    }

    if ($_SERVER["REQUEST_METHOD"] == "POST") {
        if (empty($_POST["nama"])) {
            $error_nama = "nama tidak boleh kosong";
        } else {
            if (!preg_match("/^[a-zA-Z ]*$/", $_POST["nama"])) {
                $error_nama = "Inputan hanya boleh huruf dan spasi";
            } else {
                $nama = cek_input($_POST["nama"]);
            }
        }
        if (empty($_POST["jk"])) {
            $error_jk = "Jenis kelamin harus di klik salah satu";
        } else {
            $jk = cek_input($_POST["jk"]);
        }
        if (empty($_POST['nisn'])) {
            $error_nisn = "NISN tidak oleh kosong";
        } else {
            if (!is_numeric($_POST["nisn"])) {
                $error_nisn = "NISN hanya diperbolehkan angka";
            } else {
                $nisn = cek_input($_POST["nisn"]);
            }
        }
        if (empty($_POST['nik'])) {
            $error_nik = "NIK tidak oleh kosong";
        } else {
            if (!is_numeric($_POST["nik"])) {
                $error_nik = "NIK hanya diperbolehkan angka";
            } else {
                $nik = cek_input($_POST["nik"]);
            }
        }
        if (empty($_POST["telahir"])) {
            $error_telahir = "Tempat lahir tidak boleh kosong";
        } else {
            if (!preg_match("/^[a-zA-Z ]*$/", $_POST["telahir"])) {
                $error_telahir = "Inputan hanya boleh huruf dan spasi";
            } else {
                $telahir = cek_input($_POST["telahir"]);
            }
        }
        if (empty($_POST['talahir'])) {
            $error_talahir = "Tanggal lahir tidak oleh kosong";
        } else {
            $talahir = cek_input($_POST["talahir"]);
        }
        if (empty($_POST['noakta'])) {
            $error_noakta = "nomer akta tidak oleh kosong";
        } else {
            $noakta = cek_input($_POST["noakta"]);
        }
        if (empty($_POST["agama"])) {
            $error_agama = "Agama harus dipilih";
        } else {
            $agama = cek_input($_POST["agama"]);
        }
        if (empty($_POST["negara"]) &&  empty($_POST["na"])) {
            $error_negara = "Klik Asing negara tidak boleh kosong";
        } else {
            if (!preg_match("/^[a-zA-Z]*$/", $_POST["na"])) {
                $error_negara = "Inputan hanya boleh huruf dan spasi";
            } else {
                if ($_POST["negara"] == "Asing") {
                    $_POST["negara"] = "";
                }
                $negara = cek_input($_POST["negara"] . $_POST["na"]);
                $na = cek_input($_POST["na"]);
            }
        }
        if (empty($_POST["kebkus"])) {
            $error_kebkus = "Isi salah satu";
        } else {
            $kebkus = cek_input($_POST["kebkus"]);
        }
        if (empty($_POST["alamat"])) {
            $error_alamat = "alamat tidak boleh kosong";
        } else {
            $alamat = cek_input($_POST["alamat"]);
        }
        if (empty($_POST['rt'])) {
            $error_rt = "rt tidak oleh kosong";
        } else {
            if (!is_numeric($_POST["rt"])) {
                $error_rt = "rt hanya diperbolehkan angka";
            } else {
                $rt = cek_input($_POST["rt"]);
            }
        }
        if (empty($_POST['rw'])) {
            $error_rw = "rw tidak oleh kosong";
        } else {
            if (!is_numeric($_POST["rw"])) {
                $error_rw = "rw hanya diperbolehkan angka";
            } else {
                $rw = cek_input($_POST["rw"]);
            }
        }
        if (empty($_POST["dusun"])) {
            $error_dusun = "dusun tidak boleh kosong";
        } else {
            if (!preg_match("/^[a-zA-Z ]*$/", $_POST["dusun"])) {
                $error_dusun = "Inputan hanya boleh huruf dan spasi";
            } else {
                $dusun = cek_input($_POST["dusun"]);
            }
        }
        if (empty($_POST["desa"])) {
            $error_desa = "desa tidak boleh kosong";
        } else {
            if (!preg_match("/^[a-zA-Z ]*$/", $_POST["desa"])) {
                $error_desa = "Inputan hanya boleh huruf dan spasi";
            } else {
                $desa = cek_input($_POST["desa"]);
            }
        }
        if (empty($_POST["kec"])) {
            $error_kec = "Kecamatan tidak boleh kosong";
        } else {
            if (!preg_match("/^[a-zA-Z ]*$/", $_POST["kec"])) {
                $error_kec = "Inputan hanya boleh huruf dan spasi";
            } else {
                $kec = cek_input($_POST["kec"]);
            }
        }
        if (empty($_POST['kp'])) {
            $error_kp = "Kode Pos tidak oleh kosong";
        } else {
            if (!is_numeric($_POST["kp"])) {
                $error_kp = "Kode Pos hanya diperbolehkan angka";
            } else {
                $kp = cek_input($_POST["kp"]);
            }
        }
        if (empty($_POST["lintang"])) {
            $error_lintang = "lintang tidak boleh kosong";
        } else {
            $lintang = cek_input($_POST["lintang"]);
        }
        if (empty($_POST["bujur"])) {
            $error_bujur = "bujur tidak boleh kosong";
        } else {
            $bujur = cek_input($_POST["bujur"]);
        }
        if (empty($_POST["temting"])) {
            $error_temting = "temting tidak boleh kosong";
        } else {
            $temting = cek_input($_POST["temting"]);
        }
        if (empty($_POST["modtrans"])) {
            $error_modtrans = "modal transportasi tidak boleh kosong";
        } else {
            $modtrans = cek_input($_POST["modtrans"]);
        }
        if (!empty($_POST['kks'])) {
            $kks = cek_input($_POST["kks"]);
        }
        if (empty($_POST['anakke'])) {
            $error_anakke = "anak ke tidak oleh kosong";
        } else {
            if (!is_numeric($_POST["anakke"])) {
                $error_anakke = "anak ke hanya diperbolehkan angka";
            } else {
                $anakke = cek_input($_POST["anakke"]);
            }
        }
        if (empty($_POST["pekps"]) && empty($_POST["kps"])) {
            $error_pekps = "Klik salah satu";
        } else {
            $pekps = cek_input($_POST["pekps"]);
            if (!empty($_POST["kps"])) {
                $kps = cek_input($_POST["kps"]);
            }
        }
    }
    function cek_input($data)
    {
        $data = trim($data);
        $data = stripslashes($data);
        $data = htmlspecialchars($data);
        return $data;
    }

    ?>
    <div class="container">
        <div class="row">
            <div class="col-md-10">
                <div class="card">
                    <div class="text-center">
                        <div class="card-header">
                            Formulir Peserta Didik
                        </div>
                    </div>
                    <div class="form-group row">
                        <label class="col-sm-2 col-form-label">Tanggal</label>
                        <div class="col-sm-10">
                            <label class="col-sm-1 col-form-label">:</label>
                            <label for="nama" class="col-sm-5 col-form-label"><?= $waktut ?></label>
                        </div>
                    </div>
                    <div class="text-center">
                        <div class="card-header">
                            Data Pribadi
                        </div>
                    </div>
                    <div class="card-body">
                        <form method="post" action="<?= htmlspecialchars($_SERVER["PHP_SELF"]); ?>">
                            <div class="form-group row">
                                <label for="nama" class="col-sm-2 col-form-label">Nama Lengkap</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_nama != "" ? "is-invalid" : "") ?>" id="nama" name="nama" value="<?= $nama; ?>">
                                    <span class="text-danger"><?= $error_nama ?></span>
                                </div>
                            </div>
                            <fieldset class="form-group">
                                <div class="row">
                                    <legend class="col-form-label col-sm-2 pt-0">Jenis Kelamin</legend>
                                    <div class="col-sm-10">
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="jk" id="laki-laki" value="Laki-laki" <?php if (isset($jk) && $jk == "Laki-laki") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="laki-laki">
                                                Laki-laki
                                            </label>
                                        </div>
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="jk" id="perempuan" value="Perempuan" <?php if (isset($jk) && $jk == "Perempuan") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="perempuan">
                                                Perempuan
                                            </label>
                                        </div><br>
                                        <span class="text-danger"><?= $error_jk ?></span>
                                    </div>
                                </div>
                            </fieldset>
                            <div class="form-group row">
                                <label for="nisn" class="col-sm-2 col-form-label">NISN</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_nisn != "" ? "is-invalid" : "") ?>" id="nisn" name="nisn" value="<?= $nisn; ?>">
                                    <span class="text-danger"><?= $error_nisn ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="nik" class="col-sm-2 col-form-label">NIK</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_nik != "" ? "is-invalid" : "") ?>" id="nik" name="nik" value="<?= $nik; ?>">
                                    <span class="text-danger"><?= $error_nik ?></span>
                                </div>
                            </div>
                            <div class="form-group">
                                <div class="row">
                                    <label for="telahir" class="col-sm-2 col-form-label">Tempat Lahir</label>
                                    <div class="col-sm-3">
                                        <input type="text" class="form-control <?= ($error_telahir != "" ? "is-invalid" : "") ?>" id="telahir" name="telahir" value="<?= $telahir; ?>">
                                        <span class="text-danger"><?= $error_telahir; ?></span>
                                    </div>
                                    <label for="talahir" class="col-sm-2 col-form-label">Tanggal Lahir</label>
                                    <div class="col-sm-3">
                                        <input type="date" class="form-control <?= ($error_talahir != "" ? "is-invalid" : "") ?>" id="talahir" name="talahir" value="<?= $talahir; ?>" min="1979-12-31" max="2015-12-31">
                                        <span class="text-danger"><?= $error_talahir; ?></span>
                                    </div>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="nik" class="col-sm-2 col-form-label">No Registrasi Akta Lahir</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_noakta != "" ? "is-invalid" : "") ?>" id="noakta" name="noakta" value="<?= $noakta; ?>">
                                    <span class="text-danger"><?= $error_noakta ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="agama" class="col-sm-2 col-form-label">Agama</label>
                                <div class="col-sm-5">
                                    <select name="agama" id="" class="custom-select my-1 mr-sm-2">
                                        <option selected><?= $agama; ?></option>
                                        <option value="islam">Islam</option>
                                        <option value="kristen">Kristen</option>
                                        <option value="katolik">Katolik</option>
                                        <option value="hindu">Hindu</option>
                                        <option value="budha">Budha</option>
                                    </select>
                                    <span class="text-danger"><?= $error_agama ?></span>
                                </div>
                            </div>
                            <fieldset class="form-group">
                                <div class="row">
                                    <legend class="col-form-label col-sm-2 pt-0">Negara</legend>
                                    <div class="col-sm">
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="negara" id="Indonesia" value="Indonesia" <?php if (isset($negara) && $negarar == "Indonesia") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="Indonesia">
                                                Indonesia
                                            </label>
                                        </div>
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="negara" id="Asing" value="Asing" <?php if (isset($negara) && $negarar == "Asing") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="Asing">
                                                Asing
                                            </label>
                                        </div>
                                        <div class="col-sm-5">
                                            <input type="text" class="form-control <?= ($error_negara != "" ? "is-invalid" : "") ?>" id="negara" name="na" value="<?= $na; ?>" placeholder="Nama Negara">
                                        </div>
                                        <br><span class="text-danger"><?= $error_negara ?></span>
                                    </div>
                                </div>
                            </fieldset>
                            <fieldset class="form-group">
                                <div class="row">
                                    <legend class="col-form-label col-sm-2 pt-0">Kebutuhan Khusus</legend>
                                    <div class="col-sm">
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="kebkus" id="Kebutuhan-Khusus" value="Kebutuhan-Khusus" <?php if (isset($kebkus) && $kebkus == "Kebutuhan-Khusus") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="Kebutuhan-Khusus">
                                                Ya
                                            </label>
                                        </div>
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="kebkus" id="Tidak-Kebutuhan-Khusus" value="Tidak-Kebutuhan-Khusus" <?php if (isset($kebkus) && $kebkus == "Tidak-Kebutuhan-Khusus") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="Tidak-Kebutuhan-Khusus">
                                                Tidak
                                            </label>
                                        </div>
                                        <br><span class="text-danger"><?= $error_kebkus ?></span>
                                    </div>
                                </div>
                            </fieldset>
                            <div class="form-group row">
                                <label for="alamat" class="col-sm-2 col-form-label">Alamat Rumah</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_alamat != "" ? "is-invalid" : "") ?>" id="alamat" name="alamat" value="<?= $alamat; ?>">
                                    <span class="text-danger"><?= $error_alamat; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="rt" class="col-sm-2 col-form-label">RT :</label>
                                <div class="col-sm-3">
                                    <input type="number" class="form-control <?= ($error_rt != "" ? "is-invalid" : "") ?>" id="rt" name="rt" value="<?= $rt; ?>">
                                    <span class="text-danger"><?= $error_rt; ?></span>
                                </div>
                                <label for="rw" class="col-sm-2 col-form-label">RW :</label>
                                <div class="col-sm-3">
                                    <input type="number" class="form-control <?= ($error_rw != "" ? "is-invalid" : "") ?>" id="rw" name="rw" value="<?= $rw; ?>">
                                    <span class="text-danger"><?= $error_rw; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="dusun" class="col-sm-2 col-form-label">Dusun</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_dusun != "" ? "is-invalid" : "") ?>" id="dusun" name="dusun" value="<?= $dusun; ?>">
                                    <span class="text-danger"><?= $error_dusun; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="desa" class="col-sm-2 col-form-label">Desa</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_desa != "" ? "is-invalid" : "") ?>" id="desa" name="desa" value="<?= $desa; ?>">
                                    <span class="text-danger"><?= $error_desa; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="kec" class="col-sm-2 col-form-label">Kecamatan</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_kec != "" ? "is-invalid" : "") ?>" id="kec" name="kec" value="<?= $kec; ?>">
                                    <span class="text-danger"><?= $error_kec; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="kp" class="col-sm-2 col-form-label">Kode Pos</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_kp != "" ? "is-invalid" : "") ?>" id="kp" name="kp" value="<?= $kp; ?>">
                                    <span class="text-danger"><?= $error_kp ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="lintang" class="col-sm-2 col-form-label">Lintang</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_lintang != "" ? "is-invalid" : "") ?>" id="lintang" name="lintang" value="<?= $lintang; ?>">
                                    <span class="text-danger"><?= $error_lintang ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="bujur" class="col-sm-2 col-form-label">Bujur</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_bujur != "" ? "is-invalid" : "") ?>" id="bujur" name="bujur" value="<?= $bujur; ?>">
                                    <span class="text-danger"><?= $error_bujur ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="temting" class="col-sm-2 col-form-label">Tempat Tinggal</label>
                                <div class="col-sm-5">
                                    <select name="temting" id="" class="custom-select my-1 mr-sm-2">
                                        <option selected><?= $temting; ?></option>
                                        <option value="orang_tua">Orang Tua</option>
                                        <option value="kos">Kos</option>
                                        <option value="kontrakan">Kontrakan</option>
                                        <option value="wali">Wali</option>
                                    </select>
                                    <span class="text-danger"><?= $error_temting; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="modtrans" class="col-sm-2 col-form-label">Modal Transportasi</label>
                                <div class="col-sm-5">
                                    <select name="modtrans" id="" class="custom-select my-1 mr-sm-2">
                                        <option selected><?= $modtrans; ?></option>
                                        <option value="kendaraan_pribadi">Kendaraan Pribadi</option>
                                        <option value="angkutan_umum">Angkutan Umum</option>
                                        <option value="antar_jemput">Antar Jemput</option>
                                        <option value="jalan_kaki">Jalan Kaki</option>
                                    </select>
                                    <span class="text-danger"><?= $error_modtrans; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="kks" class="col-sm-2 col-form-label">Nomor KKS (Kartu Keluarga Sejahtera)</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_kks != "" ? "is-invalid" : "") ?>" id="kks" name="kks" value="<?= $kks; ?>" placeholder="Isi Jika Ada">
                                    <span class="text-danger"><?= $error_kks; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <label for="anakke" class="col-sm-2 col-form-label">Anak Ke</label>
                                <div class="col-sm-3">
                                    <input type="number" class="form-control <?= ($error_anakke != "" ? "is-invalid" : "") ?>" id="anakke" name="anakke" value="<?= $anakke; ?>" min="1">
                                    <span class="text-danger"><?= $error_anakke ?></span>
                                </div>
                            </div>
                            <fieldset class="form-group">
                                <div class="row">
                                    <legend class="col-form-label col-sm-2 pt-0">Penerima KPS/PKH</legend>
                                    <div class="col-sm">
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="pekps" id="Ya" value="Ya" <?php if (isset($pekps) && $pekps == "Ya") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="Ya">
                                                Ya
                                            </label>
                                        </div>
                                        <div class="custom-control custom-radio custom-control-inline">
                                            <input type="radio" name="pekps" id="Tidak" value="Tidak" <?php if (isset($pekps) && $pekps == "Tidak") echo "checked"; ?> class="custom-control-input">
                                            <label class="custom-control-label" for="Tidak">
                                                Tidak
                                            </label>
                                        </div>
                                    </div>
                                </div>
                            </fieldset>
                            <div class="form-group row">
                                <label for="kks" class="col-sm-2 col-form-label">No KPS/PKH</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control <?= ($error_pekps != "" ? "is-invalid" : "") ?>" id="kps" name="kps" value="<?= $kps; ?>" placeholder="Nomor KPS/PKH">
                                    <span class="text-danger"><?= $error_pekps; ?></span>
                                </div>
                            </div>
                            <div class="form-group row">
                                <div class="col-sm-10">
                                    <button type="submit" class="btn btn-primary">Sign in</button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>

</html>
<?php

if (!empty($nama) && !empty($jk) && !empty($nisn) && !empty($nik) && !empty($telahir) && !empty($talahir) && !empty($noakta) && !empty($agama)  && !empty($negara) &&  !empty($alamat) && !empty($rt) && !empty($rw) && !empty($dusun) && !empty($desa) && !empty($kec) && !empty($kp) && !empty($lintang) && !empty($bujur) && !empty($temting) && !empty($modtrans) && !empty($anakke)) {
    $input = "INSERT INTO formulir(input, nama, jk, nisn, nik, tempatlahir, tanggallahir, nomorakta, agama, negara, kebutuhankhusus, alamat, rt, rw, dusun, desa, kecamatan, kp, lintang, bujur, temting, modtrans, kks, anakke, pekps, kps) VALUES ('$waktu','$nama','$jk','$nisn','$nik','$telahir','$talahir','$noakta','$agama','$negara','$kebkus','$alamat','$rt','$rw','$dusun','$desa','$kec','$kp','$lintang','$bujur','$temting','$modtrans','$kks','$anakke','$pekps','$kps')";
    $sql = mysqli_query($conn, $input);
    if ($sql) {
        echo "sukses";


?>
        <script>
            window.alert(
                "Form sukses diinput \n\nTanggal : <?= $waktut; ?> \nJam : <?= $jam; ?>"
            )
        </script>
        <meta http-equiv="refresh" content="0; url=reportexcel/formulir.php">

<?php
    } else {
        echo "gagal diinput";
    }
}
?>