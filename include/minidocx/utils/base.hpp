
#pragma once

#include <memory>


namespace NAMESPACE
{
  class Destroyable
  {
  private:
    bool destroyed_{ false };

  protected:
    virtual void clear() {}

  public:
    virtual void destroy() { clear(); destroyed_ = true; }
    inline bool destroyed() const { return destroyed_; }
  };


  template<class P>
  class Configurable
  {
  public:
    inline void setProperties(P props) const { props_ = std::move(props); };
    inline P& properties() { return props_; };
    inline const P& properties() const { return props_; };

  protected:
    P props_;
  };


  template<class T>
  class Variant
  {
  public:
    Variant(const T type) : type_{ type } {}

    inline T type() const { return type_; };

  private:
    const T type_;
  };


  template<class T>
  class Node : public Variant<T>, public Destroyable
  {
  public:
    Node(const T type) : Variant<T>(type) {}
    virtual ~Node() = default;
  };
}
