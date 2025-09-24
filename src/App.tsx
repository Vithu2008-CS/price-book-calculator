import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Container, Typography, FormControl, InputLabel, Select, MenuItem, Button, Box, TextField, Accordion, AccordionSummary, AccordionDetails, List, ListItem, ListItemText, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Paper, CircularProgress, Alert,
  type SelectChangeEvent
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import "./App.css";

function App() {
  const [priceData, setPriceData] = useState<PriceData[]>([]);
  const [terms, setTerms] = useState<string[]>([]);
  const [tier1Cities, setTier1Cities] = useState<string[]>([]);
  const [regions, setRegions] = useState<(string | null)[]>([]);
  const [countries, setCountries] = useState<(string | null)[]>([]);
  const [selectedRegion, setSelectedRegion] = useState('');
  const [selectedCountry, setSelectedCountry] = useState('');
  const [selectedCategory, setSelectedCategory] = useState('');
  const [selectedOption, setSelectedOption] = useState('');
  const [distance, setDistance] = useState(0);
  const [isOutOfHours, setIsOutOfHours] = useState(false);
  const [isWeekendHoliday, setIsWeekendHoliday] = useState(false);
  const [result, setResult] = useState<{ price: string; currency: string } | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const categories = {
    'Yearly Rates': ['l1_with_backfill_yearly', 'l1_without_backfill_yearly', 'l2_with_backfill_yearly', 'l2_without_backfill_yearly', 'l3_with_backfill_yearly', 'l3_without_backfill_yearly', 'l4_with_backfill_yearly', 'l4_without_backfill_yearly', 'l5_with_backfill_yearly', 'l5_without_backfill_yearly'],
    'Visit Rates': ['full_day_l1_daily', 'full_day_l2_daily', 'full_day_l3_daily', 'half_day_l1_daily', 'half_day_l2_daily', 'half_day_l3_daily'],
    'Dispatch Rates': ['dispatch_9x5x4', 'dispatch_24x7x4', 'dispatch_sbd', 'dispatch_nbd', 'dispatch_2bd', 'dispatch_3bd', 'dispatch_additional_hour'],
    'IMAC Pricing': ['imac_2bd', 'imac_3bd', 'imac_4bd'],
    'Short Term Project': ['short_term_l1_monthly', 'short_term_l2_monthly', 'short_term_l3_monthly', 'short_term_l4_monthly', 'short_term_l5_monthly'],
    'Long Term Project': ['long_term_l1_monthly', 'long_term_l2_monthly', 'long_term_l3_monthly', 'long_term_l4_monthly', 'long_term_l5_monthly']
  };

  interface PriceData {
    region: string | null;
    country: string | null;
    supplier: string | null;
    currency: string | null;
    payment_terms: string | null;
    l1_with_backfill_yearly: number | null;
    l1_without_backfill_yearly: number | null;
    l2_with_backfill_yearly: number | null;
    l2_without_backfill_yearly: number | null;
    l3_with_backfill_yearly: number | null;
    l3_without_backfill_yearly: number | null;
    l4_with_backfill_yearly: number | null;
    l4_without_backfill_yearly: number | null;
    l5_with_backfill_yearly: number | null;
    l5_without_backfill_yearly: number | null;
    full_day_l1_daily: number | null;
    full_day_l2_daily: number | null;
    full_day_l3_daily: number | null;
    half_day_l1_daily: number | null;
    half_day_l2_daily: number | null;
    half_day_l3_daily: number | null;
    dispatch_9x5x4: number | null;
    dispatch_24x7x4: number | null;
    dispatch_sbd: number | null;
    dispatch_nbd: number | null;
    dispatch_2bd: number | null;
    dispatch_3bd: number | null;
    dispatch_additional_hour: number | null;
    imac_2bd: number | null;
    imac_3bd: number | null;
    imac_4bd: number | null;
    short_term_l1_monthly: number | null;
    short_term_l2_monthly: number | null;
    short_term_l3_monthly: number | null;
    short_term_l4_monthly: number | null;
    short_term_l5_monthly: number | null;
    long_term_l1_monthly: number | null;
    long_term_l2_monthly: number | null;
    long_term_l3_monthly: number | null;
    long_term_l4_monthly: number | null;
    long_term_l5_monthly: number | null;
  }

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (event: ProgressEvent<FileReader>) => {
      try {
        const data = event.target?.result as string;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, blankrows: false }) as (string | number | null)[][];

        const extractedPriceData: PriceData[] = [];
        let dataEndIndex = sheet.length;
        for (let i = 3; i < sheet.length; i++) {
          const row = sheet[i];
          if (row.length < 5 || !row[0] || typeof row[0] === 'string' && row[0].startsWith('*')) {
            dataEndIndex = i;
            break;
          }
          extractedPriceData.push({
            region: row[0] ? String(row[0]).trim() : null,
            country: row[1] ? String(row[1]).trim() : null,
            supplier: row[2] ? String(row[2]) : null,
            currency: row[3] ? String(row[3]) : null,
            payment_terms: row[4] ? String(row[4]) : null,
            l1_with_backfill_yearly: row[5] as number | null,
            l1_without_backfill_yearly: row[6] as number | null,
            l2_with_backfill_yearly: row[7] as number | null,
            l2_without_backfill_yearly: row[8] as number | null,
            l3_with_backfill_yearly: row[9] as number | null,
            l3_without_backfill_yearly: row[10] as number | null,
            l4_with_backfill_yearly: row[11] as number | null,
            l4_without_backfill_yearly: row[12] as number | null,
            l5_with_backfill_yearly: row[13] as number | null,
            l5_without_backfill_yearly: row[14] as number | null,
            full_day_l1_daily: row[15] as number | null,
            full_day_l2_daily: row[16] as number | null,
            full_day_l3_daily: row[17] as number | null,
            half_day_l1_daily: row[18] as number | null,
            half_day_l2_daily: row[19] as number | null,
            half_day_l3_daily: row[20] as number | null,
            dispatch_9x5x4: row[21] as number | null,
            dispatch_24x7x4: row[22] as number | null,
            dispatch_sbd: row[23] as number | null,
            dispatch_nbd: row[24] as number | null,
            dispatch_2bd: row[25] as number | null,
            dispatch_3bd: row[26] as number | null,
            dispatch_additional_hour: row[27] as number | null,
            imac_2bd: row[28] as number | null,
            imac_3bd: row[29] as number | null,
            imac_4bd: row[30] as number | null,
            short_term_l1_monthly: row[31] as number | null,
            short_term_l2_monthly: row[32] as number | null,
            short_term_l3_monthly: row[33] as number | null,
            short_term_l4_monthly: row[34] as number | null,
            short_term_l5_monthly: row[35] as number | null,
            long_term_l1_monthly: row[36] as number | null,
            long_term_l2_monthly: row[37] as number | null,
            long_term_l3_monthly: row[38] as number | null,
            long_term_l4_monthly: row[39] as number | null,
            long_term_l5_monthly: row[40] as number | null,
          });
        }

        const filteredPriceData = extractedPriceData.filter(p => p.region);
        setPriceData(filteredPriceData);

        const extractedTerms: string[] = [];
        for (let i = dataEndIndex; i < sheet.length; i++) {
          const row = sheet[i];
          if (row[0] && typeof row[0] === 'string' && row[0].startsWith('*')) {
            extractedTerms.push(row[0].replace(/^\*\s*/, '').trim());
          }
        }
        setTerms(extractedTerms);

        const extractedCities: string[] = [];
        let citiesStarted = false;
        for (let i = dataEndIndex; i < sheet.length; i++) {
          const row = sheet[i];
          if (row[0] && typeof row[0] === 'string' && row[0].startsWith('* Tier 1 Cities')) {
            citiesStarted = true;
            continue;
          }
          if (citiesStarted && row[0] && typeof row[0] === 'string' && !row[0].startsWith('*')) {
            extractedCities.push(row[0].trim());
          }
        }
        setTier1Cities(extractedCities);

        const uniqueRegions = [...new Set(filteredPriceData.map(p => p.region))].sort();
        setRegions(uniqueRegions);

        setLoading(false);
      } catch (err: any) {
        setError('Error parsing Excel: ' + err.message);
        setLoading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleRegionChange = (event: SelectChangeEvent<string>)=> {
    const region = event.target.value;
    setSelectedRegion(region);
    setSelectedCountry('');
    setSelectedCategory('');
    setSelectedOption('');
    if (region) {
      const filteredCountries: (string | null)[] = [...new Set(priceData.filter((p: PriceData) => p.region === region).map((p: PriceData) => p.country))].sort();
      setCountries(filteredCountries);
    } else {
      setCountries([]);
    }
  };

  const handleCalculate = () => {
    setError(null);
    const countryData = priceData.find(p => p.country === selectedCountry);
    if (!countryData || !selectedOption) {
      setError('Select all required fields.');
      return;
    }

    let price = parseFloat(String(countryData[selectedOption as keyof PriceData])) || 0;
    const currency = countryData.currency;

    if (selectedCategory.includes('Project')) price *= 1.05;

    if (isOutOfHours) price *= 1.5;
    if (isWeekendHoliday) price *= 2;

    if (distance > 50) price += (distance - 50) * 0.4;

    setResult({ price: price.toFixed(2), currency: currency || '' });
  };

  return (
    <div className="app-container">
      <Typography variant="h4" align="center" className="app-title" gutterBottom>
        TECEZE Price Book Calculator
      </Typography>
      <Button variant="contained" component="label" className="upload-button">
        Upload Excel File
        <input type="file" hidden accept=".xlsx" onChange={handleFileUpload} />
      </Button>
      {loading && <CircularProgress className="loading-spinner" />}
      {error && <Alert severity="error" className="error-message">{error}</Alert>}
      {priceData.length > 0 && (
        <Box className="form-container">
          <FormControl fullWidth className="form-control">
            <InputLabel>Select Region</InputLabel>
            <Select value={selectedRegion} onChange={handleRegionChange}>
              <MenuItem value="">--Select--</MenuItem>
              {regions.filter((r) => r !== null).map((r) => <MenuItem key={r} value={r}>{r}</MenuItem>)}
            </Select>
          </FormControl>
          <FormControl fullWidth className="form-control" disabled={!selectedRegion}>
            <InputLabel>Select Country</InputLabel>
            <Select value={selectedCountry} onChange={(e) => setSelectedCountry(e.target.value)}>
              <MenuItem value="">--Select--</MenuItem>
              {countries.filter((c) => c !== null).map((c) => <MenuItem key={c} value={c}>{c}</MenuItem>)}
            </Select>
          </FormControl>
          <FormControl fullWidth className="form-control" disabled={!selectedCountry}>
            <InputLabel>Select Category</InputLabel>
            <Select value={selectedCategory} onChange={(e) => { setSelectedCategory(e.target.value); setSelectedOption(''); }}>
              <MenuItem value="">--Select--</MenuItem>
              {Object.keys(categories).map((cat) => <MenuItem key={cat} value={cat}>{cat}</MenuItem>)}
            </Select>
          </FormControl>
          {selectedCategory && (
            <FormControl fullWidth className="form-control">
              <InputLabel>Select Option</InputLabel>
              <Select value={selectedOption} onChange={(e) => setSelectedOption(e.target.value)}>
                <MenuItem value="">--Select--</MenuItem>
                {categories[selectedCategory as keyof typeof categories]?.map((opt) => <MenuItem key={opt} value={opt}>{opt.replace(/_/g, ' ').toUpperCase()}</MenuItem>)}
              </Select>
            </FormControl>
          )}
          <TextField label="Travel Distance (km)" type="number" value={distance} onChange={(e) => setDistance(parseFloat(e.target.value) || 0)} fullWidth className="form-control" />
          <FormControl fullWidth className="form-control">
            <InputLabel>Out of Hours</InputLabel>
            <Select value={isOutOfHours} onChange={(e) => setIsOutOfHours(e.target.value === 'true')}>
              <MenuItem value="false">No</MenuItem>
              <MenuItem value="true">Yes (1.5x)</MenuItem>
            </Select>
          </FormControl>
          <FormControl fullWidth className="form-control">
            <InputLabel>Weekend/Holiday</InputLabel>
            <Select value={isWeekendHoliday} onChange={(e) => setIsWeekendHoliday(e.target.value === 'true')}>
              <MenuItem value="false">No</MenuItem>
              <MenuItem value="true">Yes (2x)</MenuItem>
            </Select>
          </FormControl>
          <Button variant="contained" color="primary" onClick={handleCalculate} disabled={!selectedOption} className="calculate-button">
            Calculate Price
          </Button>
          {result && <Typography variant="h6" align="center" color="primary" className="result-text">Final Price: {result.price} {result.currency}</Typography>}
          <Accordion className="accordion">
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography>Terms and Conditions</Typography>
            </AccordionSummary>
            <AccordionDetails>
              <List>
                {terms.map((t, i) => <ListItem key={i}><ListItemText primary={t} /></ListItem>)}
              </List>
            </AccordionDetails>
          </Accordion>
          <Accordion className="accordion">
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography>USA Tier 1 Cities</Typography>
            </AccordionSummary>
            <AccordionDetails>
              <TableContainer component={Paper} className="table-container">
                <Table>
                  <TableHead><TableRow><TableCell>City</TableCell></TableRow></TableHead>
                  <TableBody>
                    {tier1Cities.map((city, i) => <TableRow key={i}><TableCell>{city}</TableCell></TableRow>)}
                  </TableBody>
                </Table>
              </TableContainer>
            </AccordionDetails>
          </Accordion>
        </Box>
      )}
    </div>
  );
}

export default App;