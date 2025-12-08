import axios from 'axios';

export interface IFlightStatus {
  flightNumber: string;
  airline: string;
  status: string;
  departureAirport: string;
  arrivalAirport: string;
  departureScheduled: string | null;
  arrivalScheduled: string | null;
  terminalDeparture?: string | null;
  gateDeparture?: string | null;
  terminalArrival?: string | null;
  gateArrival?: string | null;
  delayMinutes?: number | null;
  lastUpdated?: string | null;
  // NEW: airport codes for mapping
  departureIata?: string | null;
  arrivalIata?: string | null;
}

export class FlightStatusService {

  // AviationStack endpoint (must support HTTPS on your plan)
  private static readonly BASE_URL: string =
    'https://api.aviationstack.com/v1/flights';

  // 🔹 1. Single flight by flight number
  public static async getFlightStatus(
    flightNumber: string,
    apiKey: string
  ): Promise<IFlightStatus | null> {

    if (!apiKey) {
      throw new Error('API key is not configured. Please set it in the web part properties.');
    }

    if (!flightNumber) {
      throw new Error('Flight number is required.');
    }

    const url: string =
      `${FlightStatusService.BASE_URL}` +
      `?access_key=${encodeURIComponent(apiKey)}` +
      `&flight_iata=${encodeURIComponent(flightNumber)}`;

    const response = await axios.get(url);

    console.log('Single flight API response', response.data);

    const flights = response.data?.data;
    if (!flights || flights.length === 0) {
      return null;
    }

    const flight = flights[0];

    return FlightStatusService.mapFlight(flight, flightNumber);
  }

  // 🔹 2. Recent flights snapshot (free plan friendly)
  public static async getDailyFlights(
    apiKey: string,
    limit: number = 50
  ): Promise<IFlightStatus[]> {

    if (!apiKey) {
      throw new Error('API key is not configured. Please set it in the web part properties.');
    }

    const url: string =
      `${FlightStatusService.BASE_URL}` +
      `?access_key=${encodeURIComponent(apiKey)}` +
      `&limit=${limit}`;

    const response = await axios.get(url);

    console.log('Recent flights snapshot API response', response.data);

    const flights = response.data?.data;
    if (!flights || flights.length === 0) {
      return [];
    }

    return flights.map((f: any) => FlightStatusService.mapFlight(f));
  }

  // 🔹 Common mapper from raw API object → IFlightStatus
  private static mapFlight(flight: any, fallbackFlightNumber?: string): IFlightStatus {
    return {
      flightNumber: flight.flight?.iata || flight.flight?.number || fallbackFlightNumber || '',
      airline: flight.airline?.name || '',
      status: flight.flight_status || '',
      departureAirport: flight.departure?.airport || '',
      arrivalAirport: flight.arrival?.airport || '',
      departureScheduled: flight.departure?.scheduled || null,
      arrivalScheduled: flight.arrival?.scheduled || null,
      terminalDeparture: flight.departure?.terminal || null,
      gateDeparture: flight.departure?.gate || null,
      terminalArrival: flight.arrival?.terminal || null,
      gateArrival: flight.arrival?.gate || null,
      delayMinutes: flight.departure?.delay ?? null,
      lastUpdated: flight.departure?.estimated
        || flight.arrival?.estimated
        || null,
      departureIata: flight.departure?.iata || null,
      arrivalIata: flight.arrival?.iata || null
    };
  }
}
