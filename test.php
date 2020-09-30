<?php

require_once 'XlsExchange.php';

(new \XlsExchange())
    ->setInputFile('order.json')
    ->setOutputFile('items2.xlsx')
//    ->setFTP('host', 'login', 'password', 'dir')
    ->export();
