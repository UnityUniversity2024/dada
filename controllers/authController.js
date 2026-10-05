import jwt from 'jsonwebtoken';

const ADMIN_USERNAME = 'admin';
const ADMIN_PASSWORD = 'password';

export const login = (req, res) => {
  const { username, password } = req.body || {};

  if (username !== ADMIN_USERNAME || password !== ADMIN_PASSWORD) {
    return res.status(401).json({ message: 'Invalid username or password.' });
  }

  const token = jwt.sign({ username }, process.env.JWT_SECRET || 'sis-export-demo', {
    expiresIn: '8h',
  });

  return res.json({
    success: true,
    token,
    user: { username },
  });
};
