const express = require('express');
const app = express();
const stripe = require('stripe')('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');

app.use(express.static('public'));
app.use(express.json());

app.post('/process-payment', async (req, res) => {
  const { token } = req.body;

  try {
    const customer = await stripe.customers.create();
    await stripe.customers.createSource(customer.id, {
      source: token
    });

    res.send('Thẻ tín dụng đã được thêm thành công vào khách hàng trong Stripe.');
  } catch (error) {
    res.status(500).send('Đã xảy ra lỗi khi thêm thẻ tín dụng vào Stripe.');
  }
});

app.listen(3000, () => {
  console.log('Server đang lắng nghe trên cổng 3000');
});
