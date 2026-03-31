"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateId = void 0;
const crypto_1 = require("crypto");
const generateId = () => {
    try {
        return (0, crypto_1.randomUUID)();
    }
    catch {
        return `id_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
    }
};
exports.generateId = generateId;
