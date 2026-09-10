export default function handler(_req, res) {
  res.status(200).json({
    ok: true,
    message: 'Lecture Hall Allocation Reporting System API is running.',
    timestamp: new Date().toISOString()
  });
}
