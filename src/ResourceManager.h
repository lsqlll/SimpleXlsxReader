#ifndef RESOURCEMANAGER_H
#define RESOURCEMANAGER_H

#include <memory>
#include <vector>

template<typename T>
class ResourceManager
{
private:
    std::vector<std::unique_ptr<T>> resources_;

public:
    T
    registerResource (std::unique_ptr<T> resource)
    {
	resources_.push_back (std::move (resource));
	return resources_.back ().get ();
    }

    void
    cleanup ()
    {
	resources_.clear ();
    }

    ~ResourceManager () { cleanup (); }
};

#endif