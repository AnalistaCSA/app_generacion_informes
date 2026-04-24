const API_URL = "https://app-generacion-informes.onrender.com/datos";

export const obtenerDatos = async () => {
    try {
        const response = await fetch(API_URL);
        const data = await response.json();

        console.log("DATOS:", data.length);

        return data;
    } catch (error) {
        console.error("ERROR:", error);
        return [];
    }
};