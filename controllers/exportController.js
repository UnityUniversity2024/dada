import { bootstrapExports, listCoffee, listSesame, createCoffeeRecord, createSesameRecord, updateCoffeeRecord, updateSesameRecord, deleteCoffeeRecord, deleteSesameRecord } from '../services/exportService.js';

export const getHealth = (req, res) => {
  res.json({
    success: true,
    message: 'SIS export backend is running.',
    db: 'PostgreSQL / local mocks',
  });
};

export const bootstrap = async (req, res) => {
  try {
    const payload = await bootstrapExports();
    res.json(payload);
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const getCoffee = async (req, res) => {
  try {
    const rows = await listCoffee();
    res.json({ success: true, data: rows });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const postCoffee = async (req, res) => {
  try {
    const record = await createCoffeeRecord(req.body);
    res.status(201).json({ success: true, data: record });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const updateCoffee = async (req, res) => {
  try {
    const record = await updateCoffeeRecord(req.params.id, req.body);
    res.json({ success: true, data: record });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const deleteCoffee = async (req, res) => {
  try {
    const record = await deleteCoffeeRecord(req.params.id);
    res.json({ success: true, data: record });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const getSesame = async (req, res) => {
  try {
    const rows = await listSesame();
    res.json({ success: true, data: rows });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const postSesame = async (req, res) => {
  try {
    const record = await createSesameRecord(req.body);
    res.status(201).json({ success: true, data: record });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const updateSesame = async (req, res) => {
  try {
    const record = await updateSesameRecord(req.params.id, req.body);
    res.json({ success: true, data: record });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};

export const deleteSesame = async (req, res) => {
  try {
    const record = await deleteSesameRecord(req.params.id);
    res.json({ success: true, data: record });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};
