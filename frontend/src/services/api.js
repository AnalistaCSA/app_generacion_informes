const API_URL = "https://five.epicollect.net/api/export/entries/csa-ups-instalacion?form_ref=fff4776480684a35b8765ec74e7c14f8_69c54ba08a99d";

export const obtenerDatos = async () => {
    try {
        const response = await fetch(API_URL);
        const data = await response.json();

        console.log("DATA REAL:", data);

        return data.data.entries || [];
    } catch (error) {
        console.error("ERROR:", error);
        return [];
    }
};