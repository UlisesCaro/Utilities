using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Utilities
{
    internal class ImplementacionServicioRefleccion : ServicioRefleccion
    {
        private readonly Dictionary<string, PropertyInfo> _propertyCache = new Dictionary<string, PropertyInfo>();
        private readonly Dictionary<string, IList<PropertyInfo>> _propertiesCache = new Dictionary<string, IList<PropertyInfo>>();
        private readonly Dictionary<string, IList<FieldInfo>> _fieldListCache = new Dictionary<string, IList<FieldInfo>>();

        protected override object ObtenerValorDePropiedadPorRutaImpl(object objeto, string ruta)
        {
            if (!(ruta.IndexOf(".") > 0)) return ObtenerValorDePropiedad(objeto, ruta);
            var elementosDeRuta = ruta.Split(new[] { '.' });
            var objetoActual = objeto;
            var valorActual = new object();
            for (var i = 0; i < elementosDeRuta.Length; i++)
            {
                var rutaActual = elementosDeRuta[i];
                valorActual = ObtenerValorDePropiedad(objetoActual, rutaActual);
                if (i != elementosDeRuta.Length - 1)
                {
                    if (valorActual == null)
                        throw new ArgumentOutOfRangeException("ruta", ruta, "La ruta entregada no es válida");
                    objetoActual = valorActual;
                }
                else return valorActual;
            }
            return valorActual;
        }

        protected override void EstablecerValorDePropiedadImpl(object objeto, object valor, string nombrePropiedad)
        {
            var pi = ObtenerPropiedad(objeto.GetType(), nombrePropiedad);
            pi.SetValue(objeto, valor, null);
        }

        protected override object ObtenerValorDePropiedadImpl(object objeto, string nombrePropiedad)
        {
            if (objeto == null) return null;
            var pi = ObtenerPropiedad(objeto.GetType(), nombrePropiedad);
            return pi.GetValue(objeto, null);
        }

        protected override Dictionary<string, object> PropiedadesADiccionarioImpl(object objeto)
        {
            if (objeto == null) return new Dictionary<string, object>();
            var propiedades = objeto.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
            return propiedades.ToDictionary(prop => prop.Name, prop => prop.GetValue(objeto, null));
        }

        protected override PropertyInfo ObtenerPropiedadDesdeRutaImpl(Type tipoRaiz, string ruta)
        {
            if (tipoRaiz == null)
                throw new ArgumentNullException("tipoRaiz");
            if (string.IsNullOrEmpty(ruta))
                throw new ArgumentNullException("ruta");

            Type propertyType = tipoRaiz;
            PropertyInfo propertyInfo = null;

            try
            {
                var elementosDeRuta = ruta.Split(new[] { '.' });
                foreach (var rutaActual in elementosDeRuta)
                {
                    propertyInfo = ObtenerPropiedad(propertyType, rutaActual);
                    propertyType = propertyInfo.PropertyType;
                }
                return propertyInfo;
            }
            catch (ArgumentOutOfRangeException e)
            {
                throw new ArgumentOutOfRangeException("ruta", ruta, "La ruta entregada no es válida: " + e.Message);
            }
        }

        protected override PropertyInfo ObtenerPropiedadImpl(Type tipo, string nombre)
        {
            var key = tipo.FullName + "." + nombre;

            var propertyInfo = ObtienePropertyInfoDeCache(tipo, nombre, key);

            if (propertyInfo == null)
                throw new ArgumentOutOfRangeException("nombre", nombre, "La propiedad entregada no es válida");

            return propertyInfo;
        }

        private PropertyInfo ObtienePropertyInfoDeCache(Type tipo, string nombre, string key)
        {
            PropertyInfo propertyInfo;
            if (key.Contains("__AnonymousType"))
                propertyInfo = ObtienePropertyInfo(tipo, nombre);
            else
            {
                if (!_propertyCache.ContainsKey(key)) _propertyCache.Add(key, ObtienePropertyInfo(tipo, nombre));

                propertyInfo = _propertyCache[key];
            }
            return propertyInfo;
        }

        private static PropertyInfo ObtienePropertyInfo(Type tipo, string nombre)
        {
            var propertyInfo = tipo.GetProperty(nombre, BindingFlags.Public | BindingFlags.Instance | BindingFlags.FlattenHierarchy);
            return propertyInfo;
        }

        protected override IList<FieldInfo> ObtenerMiembrosImpl(object objeto)
        {
            if (objeto == null)
                throw new ReflectionHelperException("No se puede obtener los miembros de un objeto nulo");

            var nombreDeTipo = objeto.GetType().FullName;
            // ReSharper disable PossibleNullReferenceException
            if (nombreDeTipo.Contains("__AnonymousType"))
                // ReSharper restore PossibleNullReferenceException
                return objeto.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (!_fieldListCache.ContainsKey(nombreDeTipo))
                _fieldListCache.Add(nombreDeTipo, objeto.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic));

            return _fieldListCache[nombreDeTipo];
        }

        protected override IList<T> ObtenerMiembrosDeTipoImpl<T>(object contenedor)
        {
            var miembros = contenedor.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (from m in miembros where m.FieldType.IsSubclassOf(typeof(T)) select m as T).ToList();
        }

        protected override IList<PropertyInfo> ObtenerPropiedadesDeTipoImpl<T>(object contenedor)
        {
            var properties = ObtenerPropiedadesImpl(contenedor);
            var resultados = (from p in properties
                              let interfaces = p.PropertyType.GetInterfaces()
                              where p.PropertyType.Equals(typeof(T))
                              select p).ToList();
            return resultados;
        }

        protected override IList<PropertyInfo> ObtenerPropiedadesImpl(object objeto)
        {
            if (objeto == null)
                throw new ReflectionHelperException("No se pueden obtener las propiedades de un objeto nulo");
            var fullName = objeto.GetType().FullName;
            if (fullName.Contains("__AnonymousType"))
                return objeto.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).ToList();

            if (!_propertiesCache.ContainsKey(fullName))
            {
                var propiedades = objeto.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).ToList();
                _propertiesCache.Add(fullName, propiedades);
            }
            return _propertiesCache[fullName];
        }

        protected override bool TieneCampo(Type tipo, string nombrePropiedad)
        {
            var propiedad = tipo.GetProperty(nombrePropiedad);
            return propiedad != null;
        }

        protected override void EstablecerValoresDePropiedadImpl(object origen, object destino, EMapeoDesde mapeoDesde = EMapeoDesde.Origen, List<string> propiedadesAIgnorar = null)
        {
            if (propiedadesAIgnorar == null)
                propiedadesAIgnorar = new List<string>();

            var props = ObtenerPropiedades(mapeoDesde.Equals(EMapeoDesde.Origen) ? origen : destino);

            foreach (var prop in props)
            {
                if (propiedadesAIgnorar.Any(p => p.Equals(prop.Name)))
                    continue;

                try
                {
                    if (!TienePropiedad(destino.GetType(), prop.Name))
                    {
                        if (propiedadesAIgnorar.Any(p => p.Equals(prop.Name)))
                            continue;
                        else
                            throw new Exception("No se encontro la propiedad " + prop.Name + " en el objeto de destino (" + destino.GetType().Name + ")");
                    }
                    else
                    {
                        var tipoOrigen = prop.PropertyType.Name;
                        var tipoDestino = ObtenerPropiedad(destino.GetType(), prop.Name).PropertyType.Name;

                        if (!tipoOrigen.Equals(tipoDestino))
                            throw new Exception("No se puede aplicar el valor de la propiedad " + origen.GetType().ToString() + "." + prop.Name + "(" + tipoOrigen + ") en " + destino.GetType().ToString() + "." + prop.Name + "(" + tipoDestino + ")");

                        var valor = ObtenerValorDePropiedad(origen, prop.Name);
                        if (tipoOrigen.Equals("String") && valor == null)
                            valor = "";

                        EstablecerValorDePropiedad(destino, valor, prop.Name);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
    }

    public abstract class ServicioRefleccion
    {
        private static ServicioRefleccion _implementacion;
        public static ServicioRefleccion Implementacion
        {
            get { return _implementacion ?? (_implementacion = new ImplementacionServicioRefleccion()); }
            set
            {
                _implementacion = value;
            }
        }

        protected abstract object ObtenerValorDePropiedadPorRutaImpl(object objeto, string ruta);
        public static object ObtenerValorDePropiedadPorRuta(object objeto, string ruta)
        {
            return Implementacion.ObtenerValorDePropiedadPorRutaImpl(objeto, ruta);
        }

        protected abstract void EstablecerValorDePropiedadImpl(object objeto, object valor, string nombrePropiedad);
        public static void EstablecerValorDePropiedad(object objeto, object valor, string nombrePropiedad)
        {
            Implementacion.EstablecerValorDePropiedadImpl(objeto, valor, nombrePropiedad);
        }

        protected abstract object ObtenerValorDePropiedadImpl(object objeto, string nombrePropiedad);
        public static object ObtenerValorDePropiedad(object objeto, string nombrePropiedad)
        {
            return Implementacion.ObtenerValorDePropiedadImpl(objeto, nombrePropiedad);
        }

        protected abstract Dictionary<string, object> PropiedadesADiccionarioImpl(object objeto);
        public static Dictionary<string, object> PropiedadesADiccionario(object objeto)
        {
            return Implementacion.PropiedadesADiccionarioImpl(objeto);
        }

        protected abstract PropertyInfo ObtenerPropiedadDesdeRutaImpl(Type tipoRaiz, string ruta);
        public static PropertyInfo ObtenerPropiedadDesdeRuta(Type tipoRaiz, string ruta)
        {
            return Implementacion.ObtenerPropiedadDesdeRutaImpl(tipoRaiz, ruta);
        }

        protected abstract PropertyInfo ObtenerPropiedadImpl(Type tipo, string nombre);
        public static PropertyInfo ObtenerPropiedad(Type tipo, string nombre)
        {
            return Implementacion.ObtenerPropiedadImpl(tipo, nombre);
        }


        protected abstract IList<FieldInfo> ObtenerMiembrosImpl(object objeto);
        public static IList<FieldInfo> ObtenerMiembros(object contenedor)
        {
            return Implementacion.ObtenerMiembrosImpl(contenedor);
        }

        public static bool TienePropiedad(Type tipo, string nombrePropiedad)
        {
            return Implementacion.TieneCampo(tipo, nombrePropiedad);
        }

        protected abstract bool TieneCampo(Type tipo, string nombrePropiedad);

        protected abstract IList<T> ObtenerMiembrosDeTipoImpl<T>(object contenedor) where T : class;
        public static IList<T> ObtenerMiembrosDeTipo<T>(object contenedor) where T : class
        {
            return Implementacion.ObtenerMiembrosDeTipoImpl<T>(contenedor);
        }

        protected abstract IList<PropertyInfo> ObtenerPropiedadesDeTipoImpl<T>(object contenedor) where T : class;
        public static IList<PropertyInfo> ObtenerPropiedadesDeTipo<T>(object contenedor) where T : class
        {
            return Implementacion.ObtenerPropiedadesDeTipoImpl<T>(contenedor);
        }

        protected abstract IList<PropertyInfo> ObtenerPropiedadesImpl(object objeto);
        public static IList<PropertyInfo> ObtenerPropiedades(object contenedor)
        {
            return Implementacion.ObtenerPropiedadesImpl(contenedor);
        }

        protected abstract void EstablecerValoresDePropiedadImpl(object origen, object destino, EMapeoDesde mapeoDesde = EMapeoDesde.Origen, List<string> propiedadesAIgnorar = null);
        public static void EstablecerValoresDePropiedad(object origen, object destino, EMapeoDesde mapeoDesde = EMapeoDesde.Origen, List<string> propiedadesAIgnorar = null)
        {
            Implementacion.EstablecerValoresDePropiedadImpl(origen, destino, mapeoDesde, propiedadesAIgnorar);
        }
    }
}
