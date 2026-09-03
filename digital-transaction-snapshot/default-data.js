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
    id: 'atm-success-rate',
    name: 'ATM SUCCESS RATE',
    icon: 'banknote',
    type: 'success-rate',
    transactionVolume: 98.7,
    totalValue: 0,
    target: 0,
    keyMessage: 'ATM success rate',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'pos-success-rate',
    name: 'POS SUCCESS RATE',
    icon: 'credit-card',
    type: 'success-rate',
    transactionVolume: 97.5,
    totalValue: 0,
    target: 0,
    keyMessage: 'POS success rate',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'p2p-success-rate',
    name: 'P2P SUCCESS RATE',
    icon: 'users',
    type: 'success-rate',
    transactionVolume: 99.1,
    totalValue: 0,
    target: 0,
    keyMessage: 'P2P success rate',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'cash-withdrawal',
    name: 'CASH WITHDRAWAL',
    icon: 'banknote',
    type: 'financial',
    transactionVolume: 306455,
    totalValue: 455114740.00,
    target: 340000,
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
    target: 12000,
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
    target: 1000000,
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
    target: 35000,
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
    target: 20000,
    keyMessage: 'Non-financial service',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'rtp',
    name: 'RTP',
    icon: 'arrow-left-right',
    type: 'financial',
    transactionVolume: 5400,
    totalValue: 1850000000.00,
    target: 6000,
    keyMessage: '',
    highlighted: false,
    highlightStyle: ''
  },
  {
    id: 'npg',
    name: 'NPG (CARD AND ONLINE)',
    icon: 'globe',
    type: 'financial',
    transactionVolume: 21203,
    totalValue: 6925000000.00,
    target: 25000,
    keyMessage: '',
    highlighted: false,
    highlightStyle: ''
  }
];

const availableIcons = [
  'banknote', 'credit-card', 'users', 'qr-code', 'landmark',
  'wallet', 'receipt', 'building', 'coins', 'piggy-bank',
  'hand-coins', 'bank', 'smartphone', 'globe', 'shield',
  'arrow-left-right', 'percent', 'target'
];

function getDefaultData() {
  return {
    report: { ...defaultReport },
    services: defaultServices.map(s => ({ ...s }))
  };
}

module.exports = { defaultReport, defaultServices, availableIcons, getDefaultData };
