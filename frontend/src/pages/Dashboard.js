import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { obtenerDatos } from "../services/api";
import tecnicos from "../data/tecnicos.json";
import "../styles/dashboardPages.css";

function Dashboard() {

    const navigate = useNavigate();
    const [datos, setDatos] = useState([]);
    const [seleccionados, setSeleccionados] =useState([]);
    const [cargando, setCargando] = useState(false);
    const [busqueda, setBusqueda] = useState("");
    const [filtroTecnico, setFiltroTecnico] = useState("")


    //se valida datos
    useEffect(() =>{
        const usuario = localStorage.getItem("usuario");

        if(!usuario){
            navigate("/", {replace: true});
        };
    }, [navigate]);

    //se traen datos de epicoollect
    useEffect(() =>{
        const cargarDatos = async () => {
            const res = await obtenerDatos();
            setDatos(res);
        };

        cargarDatos();
    }, []);

    // Cerrar Sesión
    const logout =() =>{
        localStorage.removeItem("usuario");
        navigate("/",{replace: true})
    };

    //caja de selección
    const handleSelection = (item) => {
        const id = item.ec5_uuid;

        if (!id) {
            console.error("ID inválido:", item);
            return;
        }

        setSeleccionados(prev => {
            if (prev.includes(id)) {
                return prev.filter(i => i !== id);
            } else {
                return [...prev, id];
            }
        });
    };

    //traer nombre de tecnico

    const traerIdTecnico = (id) =>{
        const tecnico = tecnicos.find(t => t.id === id);
        return tecnico ? tecnico.nombre : "No encontrado"
    }

    //boton descargar todos los archivos
    const descargarTodos = async () => {
        setCargando(true);

        try {
            const response = await fetch("https://app-generacion-informes.onrender.com/generar", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ ids: [] })
            });

            const blob = await response.blob();

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;

            const contentDisposition = response.headers.get("Content-Disposition");
            let nombreArchivo = "informes.zip";

            if (contentDisposition) {
                const match = contentDisposition.match(/filename="(.+)"/);
                if (match) {
                    nombreArchivo = match[1];
                }
            }

            a.download = nombreArchivo;
            a.click();

        } catch (error) {
            console.error(error);
        } finally {
            setCargando(false); // 🔥 clave
        }
    };

    //boton descargar seleccionados
    const descargarSeleccionados = async () => {
        setCargando(true);

        try {
            const response = await fetch("https://app-generacion-informes.onrender.com/generar", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ ids: seleccionados })
            });

            const blob = await response.blob();

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;

            const contentDisposition = response.headers.get("Content-Disposition");
            let nombreArchivo = "informes.zip";

            if (contentDisposition) {
                const match = contentDisposition.match(/filename="(.+)"/);
                if (match) {
                    nombreArchivo = match[1];
                }
            }

            a.download = nombreArchivo;
            a.click();

        } catch (error) {
            console.error(error);
        } finally {
            setCargando(false); // 🔥 clave
        }
    };

    //filtro y cuadro de busqueda
    const datosFiltrados = datos.filter(item => {
        const texto = busqueda.toLowerCase();

        const coincideBusqueda =
            item.title?.toLowerCase().includes(texto) ||
            item["2_ID_SEDE"]?.toLowerCase().includes(texto) ||
            item["5_NOMBRE_SEDE"]?.toLowerCase().includes(texto) ||
            item["69_NUMERO_DE_SERIE_D"]?.toLowerCase().includes(texto) ||
            item["6_DIRECCION"]?.toLowerCase().includes(texto);

        const nombreTecnico = traerIdTecnico(item["11_CODIGO_TECNICO"]);

        const coincideTecnico =
            filtroTecnico === "" || nombreTecnico === filtroTecnico;

        return coincideBusqueda && coincideTecnico;
    });
    

    return (
    <div class="contenedor-Dashboard">
        <div class="contenedor-boton">
            <button class="boton-cierre-sesion" onClick={logout}>Cerrar Sesión</button>
        </div>
        <h1 class="titulo-Dashboard" >INFORMES DE INSTALACIÓN</h1>
        <p class="texto-dashboard">Bienvenido al sistema para generacion de informes UPS via Epicollect</p>
        <a href="https://five.epicollect.net/project/csa-ups-instalacion/data" target="_blank" rel="noopener noreferrer" class="link_epicollect">
            Ver base en a Epicollect
        </a>
        

        <h2 class="titulo2-Dashboard">Listado de Informes</h2>

        <div className="contenedor-filtros">
            <input
                type="text"
                placeholder="Buscar por sede, s/n de ups o dirección..."
                value={busqueda}
                onChange={(e) => setBusqueda(e.target.value)}
                className="input-busqueda"
            />

            <select
                value={filtroTecnico}
                onChange={(e) => setFiltroTecnico(e.target.value)}
                className="select-filtro"
            >
                <option value="">Todos los técnicos</option>
                {tecnicos.map(t => (
                    <option key={t.id} value={t.nombre}>
                        {t.nombre}
                    </option>
                ))}
            </select>
        </div>

        <div class="contenedor-tabla">
            {datos.length === 0 ? (
                <p>Cargando datos...</p>
            ) : (
                <table border="1" cellPadding="10" style={{ marginTop: "20px", borderCollapse: "collapse" }}>
                    <thead class="casilla-titulos">
                        <tr>
                            <th class="columna-seleccion">Selección</th>
                            <th>ID Sede</th>
                            <th>Fecha</th>
                            <th>Ciudad</th>
                            <th>Departamento</th>
                            <th>Capacidad UPS</th>
                            <th>S/N de UPS</th>
                            <th class="columna-tecnico">Técnico</th>
                            <th>Nombre Sede</th>
                            <th class="columna-direccion">Dirección</th>
                        </tr>
                    </thead>

                    <tbody>
                        {datosFiltrados.map((item, index) => (
                            <tr key={index}>
                                <td>                                    
                                    <input
                                        class="checkbox"
                                        type="checkbox" 
                                        checked={seleccionados.includes(item.ec5_uuid)}
                                        onChange={() => handleSelection(item)}
                                    />                                    
                                </td>
                                <td>{item.title}</td>
                                <td>{item.created_at.split("T")[0]}</td>
                                <td>{item["7_CIUDAD"]}</td>
                                <td>{item["8_DEPARTAMENTO"]}</td>
                                <td>{item["66_CAPACIDAD_UPS_KVA"]+" KVA"}</td>
                                <td>{item["69_NUMERO_DE_SERIE_D"]}</td>
                                <td>{traerIdTecnico(item["11_CODIGO_TECNICO"])}</td>
                                <td>{item["5_NOMBRE_SEDE"]}</td>
                                <td>{item["6_DIRECCION"]}</td>
                            </tr>                        
                        ))}
                    </tbody>
                </table>
            )}
        </div>
        <div class="contenedor-botones-descarga">
            <button class="boton-descarga" onClick={descargarTodos} disabled={cargando}>
                {cargando ? "Generando..." : "Descargar Todos"}
            </button>

            <button class="boton-descarga" onClick={descargarSeleccionados} disabled={cargando}>
                {cargando ? "Generando..." : "Descargar Seleccionados"}
            </button>
        </div>
        {cargando && (
            <p class="texto-carga">
                ⏳ Generando y descargando informes, por favor espera...
            </p>
        )}
    </div>
    );
};

export default Dashboard;