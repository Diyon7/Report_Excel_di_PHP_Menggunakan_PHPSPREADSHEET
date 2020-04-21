<?php

$server = "localhost";
$username = "root";
$password = "";
$db = "p_formulirpesertadidik";

$conn = mysqli_connect($server, $username, $password, $db);

if (!$conn) {
    die(mysqli_connect_error());
}
