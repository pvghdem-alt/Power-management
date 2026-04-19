export interface Reading {
  id?: string;
  date: string;
  meter: string;
  value: number;
  adjustment: number;
  usage: number;
  ts: number;
}

export interface BuildingConfig {
  v: string;
  l: string;
}

export const BLD_CONFIG: BuildingConfig[] = [
  { v: "CALC_急診", l: "急診大樓" },
  { v: "CALC_AB", l: "AB棟" },
  { v: "03.職務宿舍", l: "職務宿舍" },
  { v: "04.門診大樓", l: "門診大樓" },
  { v: "05.行政大樓", l: "行政大樓" },
  { v: "06.復健大樓", l: "復健大樓" },
  { v: "07.役男宿舍", l: "役男宿舍" },
  { v: "08.七病房", l: "七病房" },
  { v: "09.動力中心", l: "動力中心" },
  { v: "10.水塔", l: "水塔" },
  { v: "11.汙水處理廠", l: "汙水處理廠" },
  { v: "13.精神科大樓", l: "精神科大樓" },
  { v: "15.松柏園", l: "松柏園" },
  { v: "16.廢棄物處理廠", l: "廢棄物處理廠" }
];

export const METERS = [
  "01.急診大樓─總盤",
  "02.急診大樓─後半段",
  "03.職務宿舍",
  "04.門診大樓",
  "05.行政大樓",
  "06.復健大樓",
  "07.役男宿舍",
  "08.七病房",
  "09.動力中心",
  "10.水塔",
  "11.汙水處理廠",
  "12.廚房",
  "13.精神科大樓",
  "14.AB棟",
  "15.松柏園",
  "16.廢棄物處理廠",
  "17.懷遠堂"
];
