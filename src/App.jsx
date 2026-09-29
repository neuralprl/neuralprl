import React, { useState, useEffect, useMemo } from 'react';
import { 
  Building2, 
  FileText, 
  Users, 
  AlertTriangle, 
  CheckCircle2, 
  Search, 
  ExternalLink, 
  LogOut, 
  Plus, 
  FolderOpen,
  ShieldCheck,
  FileSpreadsheet,
  Trash2,
  Link as LinkIcon,
  Upload
} from 'lucide-react';
import * as XLSX from 'xlsx';

const INITIAL_USERS = [
  { id: '1', email: 'neuralprl', code: 'Neuralprl@', name: 'Superadministrador', role: 'superadmin', assignedCentres: ['ALL'] }
];

const INITIAL_GENERAL_DOCS = [
  { id: 'gd1', title: 'Procedimiento General de Evacuación v2', category: 'Procedimientos', link: 'https://sharepoint.com/doc1', date: '2026-01-15' },
  { id: 'gd2', title: 'Protocolo de Actuación Accidentes Laborales', category: 'Protocolos', link: 'https://sharepoint.com/doc2', date: '2026-02-01' },
];

const INITIAL_ENTERPRISES = [
  {
    id: 'ent_1',
    name: 'Neural Prevención S.L.',
    zone: 'Nacional',
    users: ['neuralprl'],
    centres: [
      {
        id: 'c1',
        name: 'Centro Neural Madrid - Castellana',
        docs: { er: [], ir: [], pap: [], epis: [], me: [] }
      }
    ]
  }
];

const DOC_CATEGORIES = [
  { key: 'er', label: 'ER', fullName: 'Evaluación de Riesgos' },
  { key: 'ir', label: 'IR', fullName: 'Información de Riesgos' },
  { key: 'pap', label: 'PAP', fullName: 'Planificación Actividad Preventiva' },
  { key: 'epis', label: 'EPIs', fullName: 'Equipos de Protección Individual' },
  { key: 'me', label: 'ME', fullName: 'Medidas de Emergencia' }
];

const extractFileNameFromUrl = (url) => {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    const filename = parsed.pathname.split('/').pop();
    if (filename && filename.length > 0) return decodeURIComponent(filename);
  } catch (e) {
    const parts = url.split('/');
    const last = parts.pop() || parts.pop();
    if (last) return decodeURIComponent(last.split('?')[0]);
  }
  return 'Documento SharePoint';
};

export default function App() {
  const [currentUser, setCurrentUser] = useState(() => {
    const saved = localStorage.getItem('neural_current_user');
    return saved ? JSON.parse(saved) : null;
  });

  const [loginEmail, setLoginEmail] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');

  const [users, setUsers] = useState(() => {
    const saved = localStorage.getItem('neural_users');
    return saved ? JSON.parse(saved) : INITIAL_USERS;
  });

  const [enterprises, setEnterprises] = useState(() => {
    const saved = localStorage.getItem('neural_enterprises');
    return saved ? JSON.parse(saved) : INITIAL_ENTERPRISES;
  });

  const [generalDocs, setGeneralDocs] = useState(() => {
    const saved = localStorage.getItem('neural_general_docs');
    return saved ? JSON.parse(saved) : INITIAL_GENERAL_DOCS;
  });

  const [activeTab, setActiveTab] = useState('enterprises');
  const [searchTerm, setSearchTerm] = useState('');

  const [editDocModal, setEditDocModal] = useState({ 
    open: false, 
    enterpriseId: null, 
    centreId: null, 
    categoryKey: null, 
    categoryLabel: '' 
  });
  const [newLinkUrl, setNewLinkUrl] = useState('');
  const [newLinkName, setNewLinkName] = useState('');

  // Guardar automáticamente en localStorage cada vez que cambien los datos
  useEffect(() => {
    localStorage.setItem('neural_enterprises', JSON.stringify(enterprises));
  }, [enterprises]);

  useEffect(() => {
    localStorage.setItem('neural_users', JSON.stringify(users));
  }, [users]);

  useEffect(() => {
    if (currentUser) {
      localStorage.setItem('neural_current_user', JSON.stringify(currentUser));
    } else {
      localStorage.removeItem('neural_current_user');
    }
  }, [currentUser]);

  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');
    const user = users.find(
      u => u.email.trim() === loginEmail.trim() && u.code.trim() === loginCode.trim()
    );
    if (user) setCurrentUser(user);
    else setLoginError('Usuario o contraseña incorrectos.');
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem('neural_current_user');
    setLoginEmail('');
    setLoginCode('');
  };

  const accessibleEnterprises = useMemo(() => {
    if (!currentUser) return [];
    if (currentUser.role === 'superadmin') return enterprises;
    return enterprises.filter(ent => ent.users.includes(currentUser.email));
  }, [currentUser, enterprises]);

  // Lector de Excel ultra robusto para detectar columnas y agrupar perfectamente
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        const enterpriseMap = {};

        data.forEach((row, index) => {
          // Buscar claves de forma flexible (ignorando mayúsculas, minúsculas y espacios)
          const keys = Object.keys(row);
          const empresaKey = keys.find(k => k.trim().toUpperCase() === 'EMPRESA');
          const centroKey = keys.find(k => {
            const clean = k.trim().toLowerCase();
            return clean === 'centro trabajo' || clean === 'centro_trabajo' || clean === 'centro trabajo ' || clean === 'title';
          });

          const empresaName = empresaKey ? row[empresaKey]?.toString().trim() : 'Sin Empresa Asignada';
          const centroName = centroKey ? row[centroKey]?.toString().trim() : `Centro ${index + 1}`;

          if (!empresaName) return;

          if (!enterpriseMap[empresaName]) {
            enterpriseMap[empresaName] = {
              id: `ent_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
              name: empresaName,
              zone: 'General',
              users: ['neuralprl'],
              centres: []
            };
          }

          enterpriseMap[empresaName].centres.push({
            id: `c_${Date.now()}_${index}`,
            name: centroName,
            docs: { er: [], ir: [], pap: [], epis: [], me: [] }
          });
        });

        const newEnterprisesList = Object.values(enterpriseMap);
        if (newEnterprisesList.length > 0) {
          setEnterprises(newEnterprisesList);
          alert(`¡Importación exitosa! Se han agrupado ${newEnterprisesList.length} empresas correctamente.`);
        } else {
          alert('No se han encontrado columnas válidas de EMPRESA o Centro trabajo en el archivo.');
        }
      } catch (err) {
        console.error(err);
        alert('Hubo un error al leer el archivo Excel.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleUrlChange = (url) => {
    setNewLinkUrl(url);
    if (url.trim() && !newLinkName) {
      setNewLinkName(extractFileNameFromUrl(url));
    }
  };

  const handleAddLink = (e) => {
    e.preventDefault();
    if (!newLinkUrl.trim() || !editDocModal.enterpriseId || !editDocModal.centreId || !editDocModal.categoryKey) return;

    const fileName = newLinkName.trim() || extractFileNameFromUrl(newLinkUrl);
    const newDocObj = {
      id: `doc_${Date.now()}`,
      name: fileName,
      link: newLinkUrl.trim()
    };

    setEnterprises(prevEnts => prevEnts.map(ent => {
      if (ent.id === editDocModal.enterpriseId) {
        return {
          ...ent,
          centres: ent.centres.map(centre => {
            if (centre.id === editDocModal.centreId) {
              const currentList = Array.isArray(centre.docs[editDocModal.categoryKey]) 
                ? centre.docs[editDocModal.categoryKey] 
                : [];
              return {
                ...centre,
                docs: {
                  ...centre.docs,
                  [editDocModal.categoryKey]: [...currentList, newDocObj]
                }
              };
            }
            return centre;
          })
        };
      }
      return ent;
    }));

    setNewLinkUrl('');
    setNewLinkName('');
  };

  const handleDeleteLink = (docId) => {
    setEnterprises(prevEnts => prevEnts.map(ent => {
      if (ent.id === editDocModal.enterpriseId) {
        return {
          ...ent,
          centres: ent.centres.map(centre => {
            if (centre.id === editDocModal.centreId) {
              const currentList = Array.isArray(centre.docs[editDocModal.categoryKey]) 
                ? centre.docs[editDocModal.categoryKey] 
                : [];
              return {
                ...centre,
                docs: {
                  ...centre.docs,
                  [editDocModal.categoryKey]: currentList.filter(d => d.id !== docId)
                }
              };
            }
            return centre;
          })
        };
      }
      return ent;
    }));
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md w-full space-y-6">
          <div className="text-center space-y-2">
            <div className="bg-blue-600 text-white w-12 h-12 rounded-lg flex items-center justify-center mx-auto shadow-lg">
              <ShieldCheck className="w-8 h-8" />
            </div>
            <h1 className="text-2xl font-bold text-slate-800">Plataforma PRL</h1>
            <p className="text-sm text-slate-500">Gestión Documental y Prevención de Riesgos</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Usuario</label>
              <input 
                type="text" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                placeholder="Nombre de usuario"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Contraseña</label>
              <input 
                type="password" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                placeholder="Tu contraseña"
                value={loginCode}
                onChange={(e) => setLoginCode(e.target.value)}
              />
            </div>

            {loginError && (
              <div className="p-3 bg-red-50 border border-red-200 text-red-600 text-xs rounded-lg flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <button 
              type="submit" 
              className="w-full py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition shadow-md"
            >
              Iniciar Sesión
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans">
      <header className="bg-slate-800 text-white shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <ShieldCheck className="w-7 h-7 text-blue-400" />
            <div>
              <h1 className="font-bold text-lg leading-tight">Neural PRL</h1>
              <span className="text-xs text-slate-400">Panel de Control de Prevención</span>
            </div>
          </div>

          <div className="flex items-center space-x-4">
            <label className="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-2 rounded-lg text-xs font-semibold flex items-center gap-2 transition shadow-lg border border-blue-500">
              <Upload className="w-4 h-4" />
              <span>Importar Excel CENTROS NEURAL</span>
              <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" />
            </label>

            <div className="text-right hidden sm:block border-l border-slate-700 pl-4">
              <p className="text-sm font-medium">{currentUser.name}</p>
              <span className="text-xs bg-blue-900 text-blue-200 px-2 py-0.5 rounded capitalize">
                {currentUser.role}
              </span>
            </div>
            <button 
              onClick={handleLogout} 
              className="p-2 text-slate-300 hover:text-white hover:bg-slate-700 rounded-lg transition"
              title="Cerrar Sesión"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <nav className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 flex space-x-8">
          <button 
            onClick={() => setActiveTab('enterprises')}
            className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'enterprises' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
          >
            <Building2 className="w-4 h-4" />
            <span>Empresas y Centros de Trabajo</span>
          </button>

          <button 
            onClick={() => setActiveTab('general')}
            className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'general' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
          >
            <FolderOpen className="w-4 h-4" />
            <span>Documentación General</span>
          </button>

          {currentUser.role === 'superadmin' && (
            <button 
              onClick={() => setActiveTab('users')}
              className={`py-4 px-2 font-medium text-sm border-b-2 flex items-center space-x-2 ${activeTab === 'users' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              <Users className="w-4 h-4" />
              <span>Gestión de Perfiles</span>
            </button>
          )}
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 py-8 flex-1 w-full">
        {activeTab === 'enterprises' && (
          <div className="space-y-8">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Listado por Empresa y Centros</h2>
                <p className="text-sm text-slate-500">Visualización vertical de centros con estado documental (ER, IR, PAP, EPIs, ME).</p>
              </div>

              <div className="relative w-full sm:w-64">
                <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                <input 
                  type="text"
                  placeholder="Buscar centro o empresa..."
                  className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>
            </div>

            <div className="space-y-6">
              {accessibleEnterprises.map((ent) => {
                const filteredCentres = ent.centres.filter(c => 
                  c.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
                  ent.name.toLowerCase().includes(searchTerm.toLowerCase())
                );

                if (filteredCentres.length === 0 && searchTerm) return null;

                return (
                  <div key={ent.id} className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                    <div className="bg-slate-800 text-white px-6 py-4 flex justify-between items-center">
                      <div className="flex items-center space-x-3">
                        <Building2 className="w-5 h-5 text-blue-400" />
                        <div>
                          <h3 className="font-bold text-lg">{ent.name}</h3>
                          <span className="text-xs text-slate-300">Empresa / Grupo</span>
                        </div>
                      </div>
                      <span className="text-xs bg-slate-700 px-3 py-1 rounded-full text-blue-200 font-medium">
                        {ent.centres.length} Centros registrados
                      </span>
                    </div>

                    <div className="divide-y divide-slate-100 p-6 space-y-6">
                      {filteredCentres.map((centre) => (
                        <div key={centre.id} className="pt-6 first:pt-0 space-y-4">
                          <div className="flex justify-between items-center">
                            <h4 className="font-bold text-slate-800 text-base flex items-center gap-2">
                              <span className="w-2 h-2 rounded-full bg-blue-600"></span>
                              {centre.name}
                            </h4>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
                            {DOC_CATEGORIES.map(({ key, label, fullName }) => {
                              const docList = Array.isArray(centre.docs[key]) ? centre.docs[key] : [];
                              const isPresent = docList.length > 0;

                              return (
                                <div key={key} className="bg-slate-50 border border-slate-200 rounded-lg p-3 flex flex-col justify-between space-y-3">
                                  <div>
                                    <div className="flex justify-between items-center mb-1">
                                      <span className="font-black text-xs text-blue-600 bg-blue-50 px-2 py-0.5 rounded border border-blue-100" title={fullName}>
                                        {label}
                                      </span>
                                      {isPresent ? (
                                        <span className="text-[10px] font-bold px-1.5 py-0.5 bg-emerald-100 text-emerald-700 rounded-full flex items-center gap-0.5">
                                          <CheckCircle2 className="w-3 h-3" /> {docList.length}
                                        </span>
                                      ) : (
                                        <span className="text-[10px] font-bold px-1.5 py-0.5 bg-rose-100 text-rose-700 rounded-full">
                                          Pendiente
                                        </span>
                                      )}
                                    </div>

                                    {isPresent ? (
                                      <div className="space-y-1 mt-2 max-h-24 overflow-y-auto">
                                        {docList.map(doc => (
                                          <a 
                                            key={doc.id}
                                            href={doc.link} 
                                            target="_blank" 
                                            rel="noreferrer"
                                            className="text-[11px] text-slate-700 hover:text-blue-600 block truncate flex items-center gap-1 bg-white p-1 rounded border border-slate-100"
                                            title={doc.name}
                                          >
                                            <FileText className="w-3 h-3 text-blue-500 shrink-0" />
                                            <span className="truncate">{doc.name}</span>
                                          </a>
                                        ))}
                                      </div>
                                    ) : (
                                      <p className="text-[11px] text-slate-400 italic mt-1">Sin archivos</p>
                                    )}
                                  </div>

                                  {currentUser.role === 'superadmin' && (
                                    <button 
                                      onClick={() => {
                                        setEditDocModal({
                                          open: true,
                                          enterpriseId: ent.id,
                                          centreId: centre.id,
                                          categoryKey: key,
                                          categoryLabel: `${centre.name} - ${fullName} (${label})`
                                        });
                                        setNewLinkUrl('');
                                        setNewLinkName('');
                                      }}
                                      className="w-full py-1.5 bg-white border border-slate-200 text-slate-600 text-[11px] font-medium rounded hover:bg-blue-50 hover:text-blue-600 hover:border-blue-200 transition flex items-center justify-center gap-1"
                                    >
                                      <Plus className="w-3 h-3" /> Gestionar
                                    </button>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {activeTab === 'general' && (
          <div className="space-y-6">
            <div>
              <h2 className="text-xl font-bold text-slate-800">Documentación General de PRL</h2>
              <p className="text-sm text-slate-500">Procedimientos, normas y protocolos comunes a toda la organización.</p>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <table className="w-full text-left text-sm text-slate-600">
                <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                  <tr>
                    <th className="px-6 py-3">Documento</th>
                    <th className="px-6 py-3">Categoría</th>
                    <th className="px-6 py-3">Fecha</th>
                    <th className="px-6 py-3 text-right">Acceso</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {generalDocs.map((doc) => (
                    <tr key={doc.id} className="hover:bg-slate-50">
                      <td className="px-6 py-4 font-medium text-slate-800">{doc.title}</td>
                      <td className="px-6 py-4">
                        <span className="px-2 py-1 bg-slate-100 text-slate-600 text-xs rounded">
                          {doc.category}
                        </span>
                      </td>
                      <td className="px-6 py-4 text-slate-400 text-xs">{doc.date}</td>
                      <td className="px-6 py-4 text-right">
                        <a 
                          href={doc.link} 
                          target="_blank" 
                          rel="noreferrer"
                          className="text-blue-600 hover:text-blue-800 text-xs font-semibold inline-flex items-center gap-1"
                        >
                          Ver archivo <ExternalLink className="w-3 h-3" />
                        </a>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === 'users' && currentUser.role === 'superadmin' && (
          <div className="space-y-6">
            <div>
              <h2 className="text-xl font-bold text-slate-800">Usuarios Registrados</h2>
              <p className="text-sm text-slate-500">Listado de perfiles y credenciales.</p>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <table className="w-full text-left text-sm text-slate-600">
                <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                  <tr>
                    <th className="px-6 py-3">Nombre</th>
                    <th className="px-6 py-3">Usuario</th>
                    <th className="px-6 py-3">Contraseña</th>
                    <th className="px-6 py-3">Rol</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {users.map((u) => (
                    <tr key={u.id} className="hover:bg-slate-50">
                      <td className="px-6 py-4 font-medium text-slate-800">{u.name}</td>
                      <td className="px-6 py-4">{u.email}</td>
                      <td className="px-6 py-4 font-mono text-xs text-slate-600">{u.code}</td>
                      <td className="px-6 py-4">
                        <span className={`px-2 py-1 text-xs font-semibold rounded ${u.role === 'superadmin' ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700'}`}>
                          {u.role}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>

      {editDocModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl shadow-xl max-w-lg w-full p-6 space-y-6 max-h-[90vh] flex flex-col">
            <div className="flex justify-between items-center border-b pb-3">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Gestionar Enlaces SharePoint</h3>
                <p className="text-xs text-slate-500">{editDocModal.categoryLabel}</p>
              </div>
              <button 
                onClick={() => setEditDocModal({ open: false, enterpriseId: null, centreId: null, categoryKey: null, categoryLabel: '' })}
                className="text-slate-400 hover:text-slate-600 text-xl font-bold"
              >
                &times;
              </button>
            </div>

            <form onSubmit={handleAddLink} className="space-y-3 bg-slate-50 p-4 rounded-lg border border-slate-200">
              <span className="text-xs font-bold text-slate-700 uppercase block">Añadir nuevo enlace</span>
              
              <div>
                <label className="block text-xs text-slate-500 mb-1">URL / Enlace de SharePoint</label>
                <div className="relative">
                  <LinkIcon className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
                  <input 
                    type="url" 
                    required
                    className="w-full pl-9 pr-3 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="https://empresa.sharepoint.com/..."
                    value={newLinkUrl}
                    onChange={(e) => handleUrlChange(e.target.value)}
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs text-slate-500 mb-1">Nombre del Archivo</label>
                <input 
                  type="text" 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                  placeholder="Ej. Evaluacion.pdf"
                  value={newLinkName}
                  onChange={(e) => setNewLinkName(e.target.value)}
                />
              </div>

              <button 
                type="submit"
                className="w-full py-2 bg-blue-600 text-white font-medium text-xs rounded-lg hover:bg-blue-700 transition shadow flex items-center justify-center gap-1"
              >
                <Plus className="w-4 h-4" />
                Guardar Enlace
              </button>
            </form>

            <div className="flex-1 overflow-y-auto space-y-2">
              <span className="text-xs font-bold text-slate-700 uppercase block">Enlaces actuales</span>
              {(() => {
                const targetEnt = enterprises.find(e => e.id === editDocModal.enterpriseId);
                const targetCentre = targetEnt?.centres.find(c => c.id === editDocModal.centreId);
                const list = targetCentre && targetCentre.docs[editDocModal.categoryKey] ? targetCentre.docs[editDocModal.categoryKey] : [];
                
                if (list.length === 0) {
                  return <p className="text-xs text-slate-400 italic">No hay enlaces guardados todavía.</p>;
                }

                return list.map(doc => (
                  <div key={doc.id} className="flex items-center justify-between p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-xs">
                    <div className="flex items-center space-x-2 truncate pr-2">
                      <FileText className="w-4 h-4 text-blue-600 shrink-0" />
                      <span className="font-medium text-slate-700 truncate">{doc.name}</span>
                    </div>
                    <button 
                      onClick={() => handleDeleteLink(doc.id)}
                      className="text-rose-500 hover:text-rose-700 p-1"
                      title="Eliminar enlace"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ));
              })()}
            </div>

            <div className="border-t pt-3 flex justify-end">
              <button
                onClick={() => setEditDocModal({ open: false, enterpriseId: null, centreId: null, categoryKey: null, categoryLabel: '' })}
                className="px-4 py-2 bg-slate-200 text-slate-700 text-xs font-medium rounded-lg hover:bg-slate-300 transition"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
