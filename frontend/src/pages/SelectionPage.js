import { useNavigate } from "react-router-dom";
import "../styles/SelectionPage.css";

function SelectionPage() {
    const navigate = useNavigate();

    const irASena = () => {
        navigate("/dashboard");
    };

    const irAColsof = () => {
        navigate("/colsof");
    };

    return (
        <div className="selection-page">
            <h1>Bienvenido a la página de selección de proyectos</h1>
            <h2>Por favor, seleccione un proyecto para continuar</h2>

            <div className="selection-buttons">
                <button className="boton sena" onClick={irASena}></button>
                <button className="boton colsof" onClick={irAColsof}></button>
            </div>
        </div>
    );
}

export default SelectionPage;