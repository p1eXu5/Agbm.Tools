using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.Helpers.Monads
{
    public struct Maybe< T > where T : class
    {
        private readonly T _instance;
        public Maybe( T instance )
        {
            _instance = instance;
        }


        public static Maybe<T> None => new Maybe< T >();

        public Maybe< TOut > Bind< TOut >( Func< T, Maybe< TOut > > func ) where TOut : class
        {
            return _instance != null ? func( _instance ) : Maybe< TOut >.None;
        }


        public Monad< TOut > Bind< TOut >( Func< T, Monad< TOut > > func ) where TOut : struct
        {
            return _instance != null ? func( _instance ) : Monad< TOut >.None;
        }

        public (Maybe< TOut >, Maybe< TOut >) Split< TOut >( Func< T, (Maybe< TOut >, Maybe< TOut >) > func ) where TOut : class
        {
            return func( _instance );
        }
    }
}
