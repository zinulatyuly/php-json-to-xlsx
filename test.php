<?php

require_once 'XlsExchange.php';

(new \XlsExchange())
    ->setInputFile('order.json')
    ->setOutputFile('items.xlsx')
//    ->setFTP('host', 'login', 'password', 'dir')
    ->export();
