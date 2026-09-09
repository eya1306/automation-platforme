"""Engines merged in from the `capgemini_smart_x` project.

Source: https://github.com/rounaaa/capgemini_smart_x

The modules here are that project's engines, kept as close to the originals as
possible so changes can still be traced back upstream. Only three kinds of edit
were made:

  * imports made relative, so the modules work as a package;
  * the tkinter front-ends dropped — the platform supplies the UI, and a server
    has no display to open a window on;
  * where a function only printed its results, it now also returns them, so the
    run panel can show counts.

`call_tree_writer` and `req_coverage` are the two modules with more than a
mechanical change; each says at the top what it was lifted out of.
"""
