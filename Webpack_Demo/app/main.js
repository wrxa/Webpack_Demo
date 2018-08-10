import React from 'react';
import { render } from 'react-dom';
import ShowSXBackgroud from './ShowSXBackgroud';
import './main.css';
// Import bootstrap css
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap';
import 'react';
import 'react-dom';
//import $ from '../assets/jquery-vendor.js';
//import 'jquery';

//require("jquery");
require("expose-loader?$!jquery");
render( < ShowSXBackgroud / >, document.getElementById('showSXBackgroudButtonText'));
//render( < ShowSXBackgroud / >, $("#showSXBackgroudButtonText"));
