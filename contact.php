<?php
session_start();
require_once 'phpmailer/PHPMailerAutoload.php';

$errors =[];

if(isset($_POST['name'],$_POST['email'],$_POST['subject'],$_POST['message'])){
    $fields=[
        'name'=>$_POST['name'],
        'email'=>$_POST['email'],
        'subject'=>$_POST['subject'],
        'message'=>$_POST['message']
    ];
    foreach($fields as $field=>$data){
        if(empty($data)){
            $errors[]='The '.$field . ' field is required.';
        }
    }
    if(empty($errors)){
        $m=new PHPMailer;
        $m->isSMTP();
        $m->SMTPAuth=true;
        $m->Host='smtp.gmail.com';
        $m->Username='mahrankhemissi1@gmail.com';//replace with your email address
        $m->Password='ma07956092';//replace with your password
        $m->SMTPSecure='ssl';
        $m->Port=465;

        $m->isHTML();
        $m->Subject = $fields['subject'];
        $m->Body='From:  '.$fields['name'].'  ('.$fields['email'].')<p>Message:'.$fields['message'].'</p>';

        $m->FromName='Contact@mahran.tn';
        $m->AddAddress('khe.mahran@gmail.com','Mahrane Khemissi');
        if ($m->send()) {
            die();
        }
    }
}else{
    $errors[]= 'Something went wrong';
}
$_SESSION['errors']=$errors;
$_SESSION['fields']=$fields;
header ('Location:index.php');