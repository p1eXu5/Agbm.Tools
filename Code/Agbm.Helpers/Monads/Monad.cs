using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.Helpers.Monads
{
    public struct Monad< T > where T : struct
    {
        private readonly T _instance;
        public Monad( T instance )
        {
            _instance = instance;
        }

        public static Monad<T> None => new Monad< T >();

        public Monad< TOut > Bind< TOut >( Func< T, Monad< TOut > > func ) where TOut : struct
        {
            return func( _instance );
        }

        public (Monad< TOut >, Monad< TOut >) Split< TOut >( Func< T, (Monad< TOut >, Monad< TOut >) > func ) where TOut : struct
        {
            return func( _instance );
        }

        public Maybe< TOut > Bind< TOut >( Func< T, Maybe< TOut > > func ) where TOut : class
        {
            return func( _instance );
        }
    }
}
