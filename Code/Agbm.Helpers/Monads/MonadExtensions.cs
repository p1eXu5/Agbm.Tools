using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.Helpers.Monads
{
    public static class MonadExtensions
    {
        public static Monad<T> Return<T>( this T instance ) where T : struct => new Monad< T >( instance );

        public static (Monad< T >, Monad< T >) Bind< T >( this (Monad< T >, Monad< T >) monads, Func< T, Monad< T > > funcA, Func< T, Monad< T > > funcB )
            where T : struct
        {
            return (monads.Item1.Bind( funcA ), monads.Item2.Bind( funcB ));
        }

        public static (Monad<T>, Monad<T>) BindParallel< T >( this (Monad<T>, Monad<T>) monads, Func< Monad<T>, Monad<T>> funcA, Func< Monad< T >, Monad< T > > funcB )
            where T : struct
        {
            return ( funcA( monads.Item1 ), funcB( monads.Item2 ));
        }
    }
}
