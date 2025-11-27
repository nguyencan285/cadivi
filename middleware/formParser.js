// middleware/formParser.js

const express = require('express');

module.exports = function (app) {
    app.use(express.urlencoded({ extended: true }));
    app.use(express.json());
};
