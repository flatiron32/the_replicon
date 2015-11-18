#!/usr/bin/env bash
sleep 480 #Wait for network, wait longer than enter time to make sure it is done
ruby replicon.rb --load --no-date --save --submit
