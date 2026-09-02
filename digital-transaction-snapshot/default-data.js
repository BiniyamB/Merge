const defaultReport = {
  title: 'DIGITAL TRANSACTION VALUE SNAPSHOT',
  organization: 'ETHSWITCH',
  brand: 'EthioPay',
  tagline1: 'Making Payment Simple and Affordable',
  tagline2: 'One Payment. Every Possibility.',
  subtitle: 'Performance Overview by Service',
  date: '31.08.26'
};

const defaultServices = [
  {
    id: 'cash-withdrawal',
    name: 'CASH WITHDRAWAL',
    icon: 'banknote',
    type: 'financial',
    transactionVolume: 306455,
    totalValue: 455114740.00,
    keyMessage: 'Lower-value transactions',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'pos-purchase',
    name: 'POS PURCHASE',
    icon: 'credit-card',
    type: 'financial',
    transactionVolume: 10189,
    totalValue: 33248484.72,
    keyMessage: 'Moderate transaction value',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'ips-p2p',
    name: 'IPS P2P',
    icon: 'users',
    type: 'financial',
    transactionVolume: 926648,
    totalValue: 4049226900.62,
    keyMessage: 'Volume leader and value driver',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'qr',
    name: 'QR',
    icon: 'qr-code',
    type: 'financial',
    transactionVolume: 30248,
    totalValue: 268061663.69,
    keyMessage: '',
    highlighted: true,
    highlightStyle: 'orange'
  },
  {
    id: 'balance-inquiry',
    name: 'BALANCE INQUIRY & MINI STATEMENT',
    icon: 'landmark',
    type: 'non-financial',
    transactionVolume: 18406,
    totalValue: 0,
    keyMessage: 'Non-financial service',
    highlighted: false,
    highlightStyle: ''
  }
];

const availableIcons = [
  'banknote', 'credit-card', 'users', 'qr-code', 'landmark',
  'wallet', 'receipt', 'building', 'coins', 'piggy-bank',
  'hand-coins', 'bank', 'smartphone', 'globe', 'shield'
];

function getDefaultData() {
  return {
    report: { ...defaultReport },
    services: defaultServices.map(s => ({ ...s }))
  };
}

module.exports = { defaultReport, defaultServices, availableIcons, getDefaultData };
